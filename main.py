#!/bin/env python3
import re, sys, os, math, numbers, math, cmath, shutil, socket, subprocess, tempfile, time

from numpy import ndarray
from enum import Enum
from typing import Self

from datetime import datetime
from enum import IntEnum
from pathlib import Path
from typing import Any, List, Optional, Self, Tuple, Union

try: import uno
except ImportError: uno = None
from io import StringIO
from contextlib import redirect_stdout


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
    LEFT, RIGHT, BLOCK, CENTER = range(4)

  class ControlCharacter(IntEnum):
    PARAGRAPH_BREAK, LINE_BREAK, HARD_HYPHEN, SOFT_HYPHEN, HARD_SPACE, APPEND_PARAGRAPH = range(6)

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

    if auto_start: self.start()

  def __enter__(self) -> Self:
    if self.doc is None: self.start()
    return self

  def __exit__(self, exc_type, exc_val, exc_tb) -> None: self.close()

  def start(self) -> None:
    if self.doc is not None: return

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
    if doc is None: return
    try:
      doc.close(True)
      return
    except Exception: pass

    try: doc.dispose()
    except Exception: pass

  def close(self) -> None:
    """
    Повністю завершує роботу з документом і soffice.
    Метод безпечний для повторного виклику.
    """
    self.close_document()

    if self.desktop is not None:
      if not self._is_external_process:
        try: self.desktop.terminate()
        except Exception: pass

    self.desktop = None
    self.ctx = None
    self.smgr = None

    process = self._process
    self._process = None

    if process is not None:
      try:
        if process.poll() is None:
          process.terminate()
          try: process.wait(timeout=2.0)
          except subprocess.TimeoutExpired:
            process.kill()
            try: process.wait(timeout=2.0)
            except subprocess.TimeoutExpired: pass
      except Exception: pass
      self._read_process_output(process)

    profile = self._user_profile_dir
    self._user_profile_dir = None

    if profile is not None:
      try: profile.cleanup()
      except Exception: pass
    self._is_external_process = False
    self._started_by_us = False

  @staticmethod
  def _find_soffice_binary() -> str:
    for cmd in ("soffice", "libreoffice"):
      path = shutil.which(cmd)
      if path: return path
    common_paths = (
      "/usr/bin/soffice",
      "/usr/lib/libreoffice/program/soffice",
      "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      r"C:\Program Files\LibreOffice\program\soffice.exe",
      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    )
    for path in common_paths:
      if os.path.isfile(path): return path
    raise LibreOfficeNotFoundError("Виконуваний файл LibreOffice (soffice) не знайдено.")
  def _get_resolver(self) -> Tuple[Any, Any]:
    try:
      local_ctx = uno.getComponentContext()
      resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_ctx,
      )
      return local_ctx, resolver
    except Exception as e: raise LibreOfficeConnectionError(f"Не вдалося створити UNO resolver: {e}") from e

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

      self.doc.storeToURL(
        file_url,
        props,
      )

      if not temp_output.exists():
        raise LibreOfficeDocumentError(f"LibreOffice повідомив про збереження, але файл не створено: {temp_output}")

      if temp_output.stat().st_size <= 0:
        raise LibreOfficeDocumentError(f"LibreOffice створив порожній файл: {temp_output}")
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

class Output:
  def __init__(self,runarg) -> None:
    self.pretty_names = {"pi": "%pi"}
    self.eps = 1e-6
    self.cplxstyle = -1
    self.office = LibreOffice() if runarg!=0 else None

  def fmt(self, v, d=3, *, forse_exp=False) -> str:
    if type(v) == str:
      return v
    if isinstance(v, complex):

      def algb(v, d):
        real_exsists = abs(v.real) > self.eps
        imag_exsists = abs(v.imag) > self.eps
        return "{}{}{}".format(
          self.fmt(v.real, d) if real_exsists else "",
          (("+" if real_exsists else "") if v.imag >= 0 else "-") if imag_exsists else "",
          ("j" + self.fmt(abs(v.imag), d)) if imag_exsists else "",
        )

      def geom(v, d):
        deg = math.degrees(cmath.phase(v))
        return "{}e^{{{}j{}°}}".format(
          self.fmt(math.sqrt(v.real**2 + v.imag**2), d),
          "" if deg >= 0 else "-",
          self.fmt(abs(deg), d),
        )

      if self.cplxstyle > 0:
        return geom(v, d)
      elif self.cplxstyle < 0:
        return algb(v, d)
      else:
        return algb(v, d) + " = " + geom(v, d)
    if isinstance(v, (int, float, numbers.Real)):
      if v == 0:
        return "0"
      v = float(v)
      if abs(v) > self.eps**-1 or forse_exp:
        before, after = "{:e}".format(v).rsplit("e")
        return "{}{}10^{{{{{}{}}}}}".format(
          self.fmt(float(before), d) if float(before) % 1 != 0 else "",
          " cdot " if float(before) % 1 != 0 else "",
          "" if float(after) > 0 else "-",
          abs(int(after)),
        )
      if abs(v) < self.eps:
        return "0"
      res = f"{v:.{d}f}".replace(".", ",").rstrip("0")
      return res[:-1] if res[-1] == "," else res
    if isinstance(v, list):
      return "[" + "``".join([self.fmt(list_value, d) for list_value in v]) + "]"
    if isinstance(v, ndarray):
      return self.fmt(v.tolist(), d)
    else:
      raise ValueError(f"Unknown type of format target{type(v)}")

  def print(self, target: list, units: str = "", format: str = "", vars: None | list[list] = None, isdef=False) -> None:
    name = target[0]
    if name in self.pretty_names.keys():
      print("raise ValueError")
      name = self.pretty_names[name]
    print(name, "= ", end="")
    if isdef:
      print(format,units)
      return
    if vars:
      for var in vars:
        n = var[0]
        if var[0] in self.pretty_names.keys():
          var[0] = self.pretty_names[n]
      print(format.format(*[v[0] for v in vars]), end=" = ")
      print(format.format(*[self.fmt(v[1]) for v in vars]), end=" = ")
    print(self.fmt(target[1]), f"{units}")

  def lp(self, target: list, units: str = "", format: str = "", vars: None | list[list] = None, isdef=False, para_break: bool = True) -> None:
    buf = StringIO()
    with redirect_stdout(buf): self.print(target, units, format, vars, isdef)
    res = buf.getvalue()
    self.office.formula(res,para_break)
  def nl(self):
    assert self.office
    self.office.paragraph("")

#@#
OPERATORS = {
  "**": "^",
  "*": "cdot",
  "/": "over",
}

def tokenise(line:str):
  m, *args = line.strip().split('#')
  tokens = [
    [m.lastgroup, m.group()]
    for m in re.compile(
      r"(?P<COMPLEX>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\s*[+-]\s*\d+(?:\.\d+)?(?:[eE][+-]?\d+)?j)"
      r"|(?P<NUM>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?j?)"
      r"|(?P<LIT>[a-zA-Z_]\w*(?:\.[a-zA-Z_]\w*)*)"
      r"|(?P<OP>\*\*|[+\-*/])"
      r"|(?P<LP>\()|(?P<RP>\))|(?P<EQ>=)|(?P<COMMA>,)|(?P<SKIP>\s+)"
    ).finditer(m)
    if m.lastgroup != "SKIP"
  ] if m != '' else []
  return tokens, args

def isexprclosed(toks):
  toks = [t[0] for t in toks]
  return toks.count("LP") - toks.count("RP")

def mainloop(targetfile):
  lbuffer = ""
  abuffer = []
  for line in targetfile:
    indentlevel = 0
    for l in line:
      if l not in ' \t\n': break
      indentlevel += 1 
    if line.strip().startswith("#\\n"): print(f'{" "*indentlevel}{"o.nl()" if runarg!=0 else "print()"}'); continue

    if line.strip().startswith(("from","import","for","while")): 
      print(line,end='')
      continue
    if line == "\n": continue
    # args: noname, nobody, nonum
    tokens,args = tokenise(lbuffer + line)

    if isexprclosed(tokens):
      lbuffer = "".join([lbuffer,line.split('#')[0]])
      abuffer += args
      continue
    print(lbuffer+line,end='')
    args = " ".join(abuffer+args)

    lbuffer = ""
    abuffer = []
    tokens_len = len(tokens)
    name = ""
    format = ""
    vars = ""
    isdef = "OP" not in [v[0] for v in tokens]
    isFuncCall = lambda: tokens_len > i + 1 and tokens[i + 1][0] == "LP"
    for i, tok in enumerate(tokens):
      if i == 0:
        name = tok[1]
      elif tok[0] == "OP":
        op = OPERATORS[tok[1]] if tok[1] in OPERATORS.keys() else tok[1]
        if tok[1] == "**" and tokens[i + 1][0] == "LP":
          assert i + 1 < tokens_len
          alive_toks = tokens[i + 1 :]
          count = 1
          operators = []
          for j, atk in enumerate(alive_toks, start=i):
            if atk[0] == "OP": operators.append(atk[1])
            if atk[0] == "LP": count += 1
            elif atk[0] == "RP": count -= 1
            if count == 0:
              if any([[operator == t for operator in operators] for t in ["+", "-", "**"]]):
                tokens[i + 1][1] = "{{"
                tokens[j][1] = "}}"
              else:
                tokens[i + 1][1] = " left("
                tokens[j][1] = " right)"
              break
        format += f" {op} "
      elif tok[0] == "LIT":
        if not isFuncCall():
          format += "{}"
          vars += f"['{tok[1]}',{tok[1]}],"
        else:
          def replacePairOfParen(i, lp="{{", rp="}}"):
            count = 0
            alive_toks = tokens[i + 1 :]
            for j, atk in enumerate(alive_toks, start=i):
              if atk[0] == "LP":
                count += 1
              elif atk[0] == "RP":
                count -= 1
              if count == 0:
                tokens[i + 1][1] = lp
                tokens[j + 1][1] = rp
                break
          match tok[1]:
            case "sqrt":
              format += "sqrt"
              replacePairOfParen(i)
            case "abs":
              replacePairOfParen(i, "left lline ", " right rline")
      elif tok[0] == "NUM":
        format += tok[1]
      elif tok[0] == "LP":
        alive_toks = tokens[i:]
        count = 0
        for j, atk in enumerate(alive_toks, start=i):
          if atk[0] == "LP":
            count += 1
          elif atk[0] == "RP":
            count -= 1
          if count == 0:
            if atk[1] in ["+", "/"]:
              tokens[i][1] = "{{"
              tokens[j][1] = "}}"
            elif tokens[i][1] == "(":
              tokens[i][1] = " left("
              tokens[j][1] = " right)"
            break
        format += tok[1]
      elif tok[0] == "RP":
        format += tok[1]
    if name in [v[1] for v in tokens[2:]]: raise NameError("Операції накшталт x = x + 1 недопустимі, через недійсність виводу")
    print(f"{' '*indentlevel}o.{'lp' if("--office" in sys.argv or "--pdf" in sys.argv) else 'print'}"
          f"(['{name}',{name}],'{args}','{format}',[{vars}],{isdef},{not '@nobreak' in args})")

def insertheader():
  with open(sys.argv[0]) as sourcefile:
    for l in sourcefile:
      if l.startswith("#@#"): break
      if l == '\n': continue
      print(l, end="")

with open(sys.argv[1]) as targetfile:
  insertheader()
  runarg = 2 if "--pdf" in sys.argv else 1 if "--office" in sys.argv else 0
  print(f"o = Output({runarg})")
  mainloop(targetfile)

  if runarg != 0:
    print("o.office.fin()")
  if runarg == 1:
    print('subprocess.run(["libreoffice", "/tmp/current.odt"], check=True)')
  elif runarg == 2:
    print('subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf",'
      '"/tmp/current.odt", "--outdir", "/tmp/"], check=True)\n'
      'subprocess.run(["xdg-open", "/tmp/current.pdf"], check=True)'
    )
