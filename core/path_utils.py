from pathlib import Path
import sys


def resource_path(relative_path):
    """
    Devuelve la ruta de un recurso tanto en desarrollo
    como cuando la aplicación está empaquetada.
    """

    if getattr(sys, "frozen", False):
        base_path = Path(sys.executable).parent
    else:
        base_path = Path(__file__).resolve().parent.parent

    return base_path / relative_path