import sys
import importlib
from pathlib import Path
from typing import Any

from PyQt6.QtWidgets import QApplication

import main_window

app: QApplication | None = None
window: main_window.MainWindow | None = None

def init() -> Any:
	"""Initialize and start the PyQt6 application with MainWindow."""
	global app, window
	app = QApplication.instance() or QApplication(sys.argv)
	window = main_window.MainWindow()
	window.show()
	return app.exec()

def reinitialize() -> Any:
	"""Close the current window and open the application again to apply the changes."""
	global app, window
	if window is not None:
		window.close()
		window.deleteLater()
		window = None

	# Reload main package modules so all pages and components are refreshed.
	package_dir = Path(__file__).parent.resolve()
	for name, module in list(sys.modules.items()):
		try:
			mod_file = getattr(module, "__file__", None)
			if not mod_file:
				continue
			mod_path = Path(mod_file).resolve()
			if package_dir in mod_path.parents or mod_path == package_dir:
				# skip this init module to avoid reloading while we're running
				if name == __name__:
					continue
				try:
					importlib.reload(module)
				except Exception:
					# ignore reload errors for individual modules
					pass
		except Exception:
			continue

	return init()


init()

