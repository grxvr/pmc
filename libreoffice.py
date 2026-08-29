import os, shutil, socket, subprocess, tempfile, time
from datetime import datetime
from enum import IntEnum
from pathlib import Path
from typing import Any, List, Optional, Self, Tuple, Union

try:
  import uno
except ImportError:
  uno = None

class LibreOfficeError(Exception):
  """Базовий виняток LibreOffice."""
class LibreOfficeNotFoundError(LibreOfficeError):
  """LibreOffice або Python UNO не знайдено."""
class LibreOfficeLaunchError(LibreOfficeError):
  """Помилка запуску LibreOffice."""
class LibreOfficeTimeoutError(LibreOfficeError):
  """Вичерпано час очікування UNO."""
class LibreOfficeConnectionError(LibreOfficeError):
  """Помилка підключення до UNO."""
class LibreOfficeDocumentError(LibreOfficeError):
  """Помилка роботи з документом."""
class LibreOfficeBackupError(LibreOfficeError):
  """Помилка створення backup."""

class LibreOffice:
  class S(IntEnum):
    LEFT = 0
    RIGHT = 1
    BLOCK = 2
    CENTER = 3

  class ControlCharacter(IntEnum):
    PARAGRAPH_BREAK = 0
    LINE_BREAK = 1
    HARD_HYPHEN = 2
    SOFT_HYPHEN = 3
    HARD_SPACE = 4
    APPEND_PARAGRAPH = 5

  WRITER_FILTER = "writer8"
  MATH_CLSID = "078B7ABA-54FC-457F-8551-6147e776a997"

  def __init__(
    self,
    host: str = "127.0.0.1",
    port: int = 2002,
    connect_timeout: float = 15.0,
    soffice_binary: Optional[str] = None,
    backup_dir: Union[str, Path] = "/tmp/libreoffice_backups",
    headless: bool = True,
    auto_start: bool = True,
  ) -> None:
    if uno is None:
      raise LibreOfficeNotFoundError("Модуль 'uno' не знайдено. Використовуйте Python, який має доступ до LibreOffice UNO.")

    self.host = host
    self.port = port
    self.connect_timeout = float(connect_timeout)
    self.soffice_binary = soffice_binary or self._find_soffice_binary()
    self.backup_dir = Path(backup_dir)
    self.headless = headless

    self._process: Optional[subprocess.Popen] = None
    self._user_profile_dir: Optional[tempfile.TemporaryDirectory] = None
    self._is_external_process = False
    self._started_by_us = False

    self._stdout = ""
    self._stderr = ""

    self.ctx: Any = None
    self.smgr: Any = None
    self.desktop: Any = None

    self.doc: Any = None
    self.text: Any = None
    self.cursor: Any = None

    self.common_style = "Standard"
    self._graphic_provider: Any = None

    if auto_start:
      self.start()

  # ---------------------------------------------------------------------------
  # Lifecycle
  # ---------------------------------------------------------------------------

  def __enter__(self) -> Self:
    if self.doc is None:
      self.start()
    return self

  def __exit__(self, exc_type, exc_val, exc_tb) -> None:
    self.close()

  def start(self) -> None:
    if self.doc is not None:
      return

    try:
      self._connect_uno()
      self._init_document()
    except Exception:
      self.close()
      raise

  def close_document(self) -> None:
    doc = self.doc

    self.doc = None
    self.text = None
    self.cursor = None
    self._graphic_provider = None

    if doc is None:
      return

    try:
      doc.close(True)
      return
    except Exception:
      pass

    try:
      doc.dispose()
    except Exception:
      pass

  def close(self) -> None:
    """
    Повністю завершує роботу з документом і soffice.

    Метод безпечний для повторного виклику.
    """

    self.close_document()

    if self.desktop is not None:
      if not self._is_external_process:
        try:
          self.desktop.terminate()
        except Exception:
          pass

    self.desktop = None
    self.ctx = None
    self.smgr = None

    process = self._process
    self._process = None

    if process is not None:
      try:
        if process.poll() is None:
          process.terminate()

          try:
            process.wait(timeout=2.0)
          except subprocess.TimeoutExpired:
            process.kill()

            try:
              process.wait(timeout=2.0)
            except subprocess.TimeoutExpired:
              pass
      except Exception:
        pass

      self._read_process_output(process)

    profile = self._user_profile_dir
    self._user_profile_dir = None

    if profile is not None:
      try:
        profile.cleanup()
      except Exception:
        pass

    self._is_external_process = False
    self._started_by_us = False

  # ---------------------------------------------------------------------------
  # LibreOffice discovery
  # ---------------------------------------------------------------------------

  @staticmethod
  def _find_soffice_binary() -> str:
    for cmd in ("soffice", "libreoffice"):
      path = shutil.which(cmd)

      if path:
        return path

    common_paths = (
      "/usr/bin/soffice",
      "/usr/lib/libreoffice/program/soffice",
      "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      r"C:\Program Files\LibreOffice\program\soffice.exe",
      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    )

    for path in common_paths:
      if os.path.isfile(path):
        return path

    raise LibreOfficeNotFoundError("Виконуваний файл LibreOffice (soffice) не знайдено.")

  # ---------------------------------------------------------------------------
  # UNO connection
  # ---------------------------------------------------------------------------

  def _get_resolver(self) -> Tuple[Any, Any]:
    try:
      local_ctx = uno.getComponentContext()

      resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_ctx,
      )

      return local_ctx, resolver

    except Exception as e:
      raise LibreOfficeConnectionError(f"Не вдалося створити UNO resolver: {e}") from e

  def _uno_url(self) -> str:
    return f"uno:socket,host={self.host},port={self.port};urp;StarOffice.ComponentContext"

  def _try_connect_once(self) -> bool:
    try:
      _, resolver = self._get_resolver()

      self.ctx = resolver.resolve(self._uno_url())
      self.smgr = self.ctx.ServiceManager

      self.desktop = self.smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop",
        self.ctx,
      )

      return self.desktop is not None

    except Exception:
      self.ctx = None
      self.smgr = None
      self.desktop = None
      return False

  def _port_available(self) -> bool:
    """
    Перевіряє, чи можна підключитися до порту.

    Це не замінює UNO connection, але дозволяє швидше
    діагностувати проблеми запуску.
    """

    try:
      with socket.create_connection(
        (self.host, self.port),
        timeout=0.25,
      ):
        return True
    except OSError:
      return False

  # ---------------------------------------------------------------------------
  # Process management
  # ---------------------------------------------------------------------------

  def _start_process(self) -> None:
    if self._process is not None:
      return

    try:
      self._user_profile_dir = tempfile.TemporaryDirectory(prefix="soffice_profile_")

      profile_path = Path(self._user_profile_dir.name).resolve()

      profile_url = uno.systemPathToFileUrl(str(profile_path))

      cmd = [
        self.soffice_binary,
        "--nocrashreport",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--nologo",
        "--norestore",
        f"--accept=socket,host={self.host},port={self.port};urp;",
        f"-env:UserInstallation={profile_url}",
      ]

      if self.headless:
        cmd.extend(
          [
            "--headless",
            "--invisible",
          ]
        )

      self._process = subprocess.Popen(
        cmd,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        close_fds=True,
      )

      self._started_by_us = True

    except OSError as e:
      self._cleanup_profile()

      raise LibreOfficeLaunchError(f"Помилка запуску LibreOffice: {e}") from e

  def _read_process_output(
    self,
    process: subprocess.Popen,
  ) -> None:
    """
    Забирає stdout/stderr після завершення процесу.

    communicate() не використовується до terminate(),
    щоб не блокувати нормальну роботу UNO.
    """

    try:
      stdout, stderr = process.communicate(timeout=0.2)

      if stdout:
        self._stdout += stdout

      if stderr:
        self._stderr += stderr

    except Exception:
      pass

  def _cleanup_profile(self) -> None:
    profile = self._user_profile_dir
    self._user_profile_dir = None

    if profile is not None:
      try:
        profile.cleanup()
      except Exception:
        pass

  def _connect_uno(self) -> None:
    # Спочатку пробуємо підключитися до вже існуючого soffice.
    if self._try_connect_once():
      self._is_external_process = True
      return

    self._start_process()

    start = time.monotonic()

    while True:
      elapsed = time.monotonic() - start

      if elapsed >= self.connect_timeout:
        self.close()

        raise LibreOfficeTimeoutError(f"Timeout ({self.connect_timeout:.1f}s) очікування UNO на {self.host}:{self.port}.")

      if self._process is not None:
        return_code = self._process.poll()

        if return_code is not None:
          self._read_process_output(self._process)

          stderr = self._stderr.strip()

          message = f"LibreOffice завершився під час запуску з кодом {return_code}."

          if stderr:
            message += f"\nstderr:\n{stderr}"

          raise LibreOfficeLaunchError(message)

      if self._try_connect_once():
        return

      time.sleep(0.1)

  # ---------------------------------------------------------------------------
  # Document initialization
  # ---------------------------------------------------------------------------

  def _init_document(self) -> None:
    try:
      self.doc = self.desktop.loadComponentFromURL(
        "private:factory/swriter",
        "_blank",
        0,
        (),
      )

      if self.doc is None:
        raise LibreOfficeDocumentError("LibreOffice не повернув документ.")

      self.text = self.doc.getText()
      self.cursor = self.text.createTextCursor()

      self.cursor.gotoEnd(False)
      self.cursor.CharHeight = 14

      self._init_common_style()

      self._graphic_provider = self.smgr.createInstanceWithContext(
        "com.sun.star.graphic.GraphicProvider",
        self.ctx,
      )

    except LibreOfficeDocumentError:
      raise

    except Exception as e:
      raise LibreOfficeDocumentError(f"Помилка створення Writer-документа: {e}") from e

  # ---------------------------------------------------------------------------
  # Backup
  # ---------------------------------------------------------------------------

  @staticmethod
  def backup_file(
    source_path: Union[str, Path],
    backup_dir: Union[str, Path] = "/tmp/libreoffice_backups",
  ) -> Optional[Path]:

    src = Path(source_path)
    dst_dir = Path(backup_dir)

    if not src.exists():
      return None

    if not src.is_file():
      raise LibreOfficeBackupError(f"Backup source не є файлом: {src}")

    try:
      dst_dir.mkdir(
        parents=True,
        exist_ok=True,
      )

      timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S_%f")

      suffix = src.suffix or ".odt"

      destination = dst_dir / f"{timestamp}{suffix}"

      counter = 1

      while destination.exists():
        destination = dst_dir / f"{timestamp}_{counter}{suffix}"

        counter += 1

      shutil.copy2(
        src,
        destination,
      )

      return destination

    except Exception as e:
      raise LibreOfficeBackupError(f"Не вдалося створити backup '{src}': {e}") from e

  # ---------------------------------------------------------------------------
  # Saving
  # ---------------------------------------------------------------------------

  def _make_temp_output(
    self,
    output: Path,
  ) -> Path:

    return output.with_name(f".{output.name}.{os.getpid()}.{time.time_ns()}.tmp")

  def fin(
    self,
    output_path: Union[str, Path] = "/tmp/current.odt",
    backup: bool = True,
    backup_dir: Union[str, Path] = "/tmp/libreoffice_backups",
    close: bool = True,
  ) -> Path:

    if self.doc is None:
      raise LibreOfficeDocumentError("Немає відкритого LibreOffice-документа.")

    output = Path(output_path).resolve()
    temp_output = self._make_temp_output(output)

    try:
      output.parent.mkdir(
        parents=True,
        exist_ok=True,
      )

      if backup and output.exists():
        self.backup_file(
          source_path=output,
          backup_dir=backup_dir,
        )

      file_url = uno.systemPathToFileUrl(str(temp_output))

      props = (
        self._prop(
          "FilterName",
          self.WRITER_FILTER,
        ),
        self._prop(
          "Overwrite",
          True,
        ),
      )

      # ВАЖЛИВО:
      #
      # storeToURL(), а не storeAsURL().
      #
      # Документ залишається прив'язаним до свого UNO lifecycle,
      # а запис виконується у тимчасовий файл.
      self.doc.storeToURL(
        file_url,
        props,
      )

      if not temp_output.exists():
        raise LibreOfficeDocumentError(f"LibreOffice повідомив про збереження, але файл не створено: {temp_output}")

      if temp_output.stat().st_size <= 0:
        raise LibreOfficeDocumentError(f"LibreOffice створив порожній файл: {temp_output}")

      # Атомарна заміна.
      #
      # Старий current.odt залишається недоторканим,
      # поки новий файл повністю не записаний.
      os.replace(
        temp_output,
        output,
      )

      result = output

    except LibreOfficeDocumentError:
      raise

    except Exception as e:
      raise LibreOfficeDocumentError(f"Помилка збереження '{output}': {e}") from e

    finally:
      try:
        if temp_output.exists():
          temp_output.unlink()
      except OSError:
        pass

      if close:
        self.close()

    return result

  # ---------------------------------------------------------------------------
  # Properties / UNO helpers
  # ---------------------------------------------------------------------------

  @staticmethod
  def _prop(
    name: str,
    value: Any,
  ) -> Any:

    prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")

    prop.Name = name
    prop.Value = value

    return prop

  @staticmethod
  def _create_line_spacing(
    height: int,
    mode: int,
  ) -> Any:

    spacing = uno.createUnoStruct("com.sun.star.style.LineSpacing")

    spacing.Mode = mode
    spacing.Height = height

    return spacing

  # ---------------------------------------------------------------------------
  # Styles
  # ---------------------------------------------------------------------------

  def _init_common_style(self) -> None:
    try:
      paragraph_styles = self.doc.getStyleFamilies().getByName("ParagraphStyles")

      for name in paragraph_styles.getElementNames():
        style = paragraph_styles.getByName(name)

        if getattr(style, "OutlineLevel", 0) == 0:
          self.common_style = name
          return

    except Exception:
      self.common_style = "Standard"

  def _get_heading_style(
    self,
    n: int,
  ) -> str:

    try:
      paragraph_styles = self.doc.getStyleFamilies().getByName("ParagraphStyles")

      for name in paragraph_styles.getElementNames():
        style = paragraph_styles.getByName(name)

        if getattr(style, "OutlineLevel", 0) == n:
          return name

    except Exception:
      pass

    return self.common_style

  # ---------------------------------------------------------------------------
  # Text
  # ---------------------------------------------------------------------------

  def header(
    self,
    s: str,
    n: int = 1,
  ) -> Self:

    try:
      self.cursor.ParaStyleName = self._get_heading_style(n)

      self.text.insertString(
        self.cursor,
        s,
        0,
      )

      self.cursor.ParaAdjust = self.S.BLOCK

      self.text.insertControlCharacter(
        self.cursor,
        self.ControlCharacter.PARAGRAPH_BREAK,
        False,
      )

      return self

    except Exception as e:
      raise LibreOfficeDocumentError(f"Помилка вставки header: {e}") from e

  # ---------------------------------------------------------------------------
  # Formula
  # ---------------------------------------------------------------------------

  def formula(
    self,
    s: str,
    para_break: bool = True,
  ) -> Self:

    try:
      formula_obj = self.doc.createInstance("com.sun.star.text.TextEmbeddedObject")

      formula_obj.setPropertyValue(
        "CLSID",
        self.MATH_CLSID,
      )

      formula_obj.setPropertyValue(
        "AnchorType",
        uno.Enum(
          "com.sun.star.text.TextContentAnchorType",
          "AS_CHARACTER",
        ),
      )

      self.text.insertTextContent(
        self.cursor,
        formula_obj,
        False,
      )

      self.cursor.ParaTopMargin = 100
      self.cursor.ParaBottomMargin = 200
      self.cursor.ParaLineSpacing = self._create_line_spacing(
        115,
        0,
      )

      embedded = None

      try:
        embedded = formula_obj.getEmbeddedObject()
      except Exception:
        embedded = getattr(
          formula_obj,
          "EmbeddedObject",
          None,
        )

      if embedded is not None:
        try:
          embedded.changeState(1)
        except Exception:
          pass

      math_component = None

      if embedded is not None:
        try:
          component = embedded.Component

          if component is not None:
            math_component = component
        except Exception:
          pass

      if math_component is None:
        try:
          model = formula_obj.Model

          if model is not None:
            math_component = model
        except Exception:
          pass

      if math_component is None:
        try:
          component = formula_obj.getComponent()

          if component is not None:
            math_component = component
        except Exception:
          pass

      if math_component is None:
        math_component = embedded

      if math_component is None:
        raise LibreOfficeDocumentError(f"Не вдалося отримати StarMath component для формули '{s}'.")

      try:
        math_component.setPropertyValue(
          "Formula",
          s,
        )
      except Exception:
        math_component.Formula = s

      if para_break:
        self.text.insertControlCharacter(
          self.cursor,
          self.ControlCharacter.PARAGRAPH_BREAK,
          False,
        )

      return self

    except LibreOfficeDocumentError:
      raise

    except Exception as e:
      raise LibreOfficeDocumentError(f"Помилка вставки formula '{s}': {e}") from e

  # ---------------------------------------------------------------------------
  # Images
  # ---------------------------------------------------------------------------

  def image(
    self,
    filename: Union[str, Path],
    zoom: float = 1.0,
  ) -> Self:

    file_path = Path(filename).resolve()

    if not file_path.is_file():
      raise LibreOfficeDocumentError(f"Зображення не знайдено: {file_path}")

    if zoom <= 0:
      raise LibreOfficeDocumentError(f"zoom повинен бути > 0, отримано {zoom}")

    try:
      file_url = uno.systemPathToFileUrl(str(file_path))

      graphic = self._graphic_provider.queryGraphic(
        (
          self._prop(
            "URL",
            file_url,
          ),
        )
      )

      image_obj = self.doc.createInstance("com.sun.star.text.TextGraphicObject")

      image_obj.setPropertyValue(
        "AnchorType",
        uno.Enum(
          "com.sun.star.text.TextContentAnchorType",
          "AS_CHARACTER",
        ),
      )

      image_obj.Graphic = graphic

      size = graphic.Size100thMM

      width = int(size.Width)
      height = int(size.Height)

      if width <= 0 or height <= 0:
        pixel_size = graphic.SizePixel

        width = int(pixel_size.Width * 2540 / 96)

        height = int(pixel_size.Height * 2540 / 96)

      page_styles = self.doc.StyleFamilies.getByName("PageStyles")

      page_style = page_styles.getByName("Standard")

      max_width = page_style.Width - page_style.LeftMargin - page_style.RightMargin

      if width > max_width and width > 0:
        scale = max_width / width

        width = int(width * scale)
        height = int(height * scale)

      image_obj.Width = int(width * zoom)
      image_obj.Height = int(height * zoom)

      self.cursor.ParaTopMargin = 0
      self.cursor.ParaBottomMargin = 0

      self.cursor.ParaLineSpacing = self._create_line_spacing(
        200,
        1,
      )

      self.cursor.ParaAdjust = self.S.CENTER

      self.text.insertTextContent(
        self.cursor,
        image_obj,
        False,
      )

      self.text.insertControlCharacter(
        self.cursor,
        self.ControlCharacter.PARAGRAPH_BREAK,
        False,
      )

      return self

    except LibreOfficeDocumentError:
      raise

    except Exception as e:
      raise LibreOfficeDocumentError(f"Помилка вставки image '{file_path}': {e}") from e

  # ---------------------------------------------------------------------------
  # Formula text
  # ---------------------------------------------------------------------------

  def insert_formula_text(
    self,
    s: str,
  ) -> None:

    token = "$"

    if s.count(token) % 2 != 0:
      raise LibreOfficeDocumentError(f"Непарна кількість '{token}' у рядку: {s}")

    parts = s.split(token)

    for i, part in enumerate(parts):
      if not part:
        continue

      if i % 2 == 0:
        self.text.insertString(
          self.cursor,
          part,
          0,
        )
      else:
        self.formula(
          part,
          para_break=False,
        )

    self.text.insertControlCharacter(
      self.cursor,
      self.ControlCharacter.PARAGRAPH_BREAK,
      False,
    )

  # ---------------------------------------------------------------------------
  # Paragraphs
  # ---------------------------------------------------------------------------

  def desc(
    self,
    s: str,
    linetype: str,
  ) -> None:

    self.insert_formula_text(s)

    self.cursor.ParaTopMargin = 0 if linetype == "image" else 200

    self.cursor.ParaBottomMargin = 200 if linetype == "image" else 0

    self.cursor.ParaLineSpacing = self._create_line_spacing(
      115,
      0,
    )

  def paragraph(
    self,
    s: str,
  ) -> None:

    self.insert_formula_text(s)

    self.cursor.ParaTopMargin = 100
    self.cursor.ParaBottomMargin = 200

    self.cursor.ParaLineSpacing = self._create_line_spacing(
      115,
      0,
    )

    self.cursor.ParaAdjust = self.S.BLOCK

  # ---------------------------------------------------------------------------
  # Tables
  # ---------------------------------------------------------------------------

  def table(
    self,
    t: List[List[str]],
  ) -> None:

    if not t or not any(t):
      return

    rows = len(t)
    cols = max(len(row) for row in t)

    if cols == 0:
      return

    try:
      table = self.doc.createInstance("com.sun.star.text.TextTable")

      table.initialize(
        rows,
        cols,
      )

      self.text.insertTextContent(
        self.cursor,
        table,
        False,
      )

      for row_index, row in enumerate(t):
        for col_index, value in enumerate(row):
          cell_name = f"{self._column_name(col_index)}{row_index + 1}"

          cell = table.getCellByName(cell_name)

          cell.setString(str(value))

          cell_cursor = cell.createTextCursor()

          cell_cursor.ParaStyleName = self.common_style

          cell_cursor.setPropertyValue(
            "ParaAdjust",
            self.S.CENTER,
          )

      table.RepeatHeadline = False
      table.HeaderRowCount = 0

    except Exception as e:
      raise LibreOfficeDocumentError(f"Помилка створення таблиці: {e}") from e

  @staticmethod
  def _column_name(index: int) -> str:
    """
    0 -> A
    1 -> B
    ...
    25 -> Z
    26 -> AA
    """

    if index < 0:
      raise ValueError("column index < 0")

    result = ""

    index += 1

    while index:
      index, remainder = divmod(
        index - 1,
        26,
      )

      result = chr(ord("A") + remainder) + result

    return result

  # ---------------------------------------------------------------------------
  # Alignment
  # ---------------------------------------------------------------------------

  def align(
    self,
    a: Union[str, int],
  ) -> None:

    try:
      par_cursor = self.text.createTextCursorByRange(self.cursor)

      par_cursor = par_cursor.queryInterface(uno.getTypeByName("com.sun.star.text.XParagraphCursor"))

      if not par_cursor:
        return

      par_cursor.gotoEndOfParagraph(False)

      if par_cursor.gotoPreviousParagraph(False):
        par_cursor.gotoStartOfParagraph(False)

      value = None

      if a in ("e", self.S.CENTER):
        value = self.S.CENTER

      elif a in ("r", self.S.RIGHT):
        value = self.S.RIGHT

      elif a in ("l", self.S.LEFT):
        value = self.S.LEFT

      elif a in ("j", self.S.BLOCK):
        value = self.S.BLOCK

      if value is not None:
        par_cursor.setPropertyValue(
          "ParaAdjust",
          value,
        )

    except Exception as e:
      raise LibreOfficeDocumentError(f"Помилка align: {e}") from e

  def allign(
    self,
    a: Union[str, int],
  ) -> None:
    self.align(a)

  # ---------------------------------------------------------------------------
  # Diagnostics
  # ---------------------------------------------------------------------------

  @property
  def process_stdout(self) -> str:
    return self._stdout

  @property
  def process_stderr(self) -> str:
    return self._stderr

  @property
  def profile_path(self) -> Optional[Path]:
    if self._user_profile_dir is None:
      return None

    return Path(self._user_profile_dir.name)

  @property
  def running(self) -> bool:
    return self._process is not None and self._process.poll() is None


lo = LibreOffice()

lo.formula("%pi = 3,14")
lo.fin("/tmp/current.odt")
lo.close()

subprocess.run(
  [
    "libreoffice",
    "--headless",
    "--convert-to",
    "pdf",
    "/tmp/current.odt",
    "--outdir",
    "/tmp/",
  ],
  check=True,
)

subprocess.run(
  ["xdg-open", "/tmp/current.pdf"],
  check=True,
)
