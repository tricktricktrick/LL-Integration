import mobase

from .plugin import (
    LoversLabInstallObserver,
    LoversLabMenuTool,
)


def createPlugins() -> list[mobase.IPlugin]:
    return [
        LoversLabMenuTool(),
        LoversLabInstallObserver(),
    ]
