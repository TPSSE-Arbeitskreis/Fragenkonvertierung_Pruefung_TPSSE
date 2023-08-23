from pathlib import Path
from datetime import datetime
import click
import sys
import yaml
from typing import Any, IO
from rich import print
import os
import json
# from .myloadenv import load_env


class Loader(yaml.SafeLoader):
    """YAML Loader with `!include` constructor."""

    def __init__(self, stream: IO) -> None:
        """Initialise Loader."""

        try:
            self._root = os.path.split(stream.name)[0]
        except AttributeError:
            self._root = os.path.curdir

        super().__init__(stream)


def construct_include(loader: Loader, node: yaml.Node) -> Any:
    """Include file referenced at node."""

    filename = os.path.abspath(os.path.join(loader._root, loader.construct_scalar(node)))
    extension = os.path.splitext(filename)[1].lstrip('.')

    with open(filename, 'r', encoding='utf8') as f:
        if extension in ('yaml', 'yml'):
            return yaml.load(f, Loader)
        elif extension in ('json',):
            return json.load(f)
        else:
            return ''.join(f.readlines())


class FragenConv:
    
    @staticmethod
    def _check_dir(path_to_check: str, basedir: str, typ: str) -> Path:
        path_to_check = Path(path_to_check)
        if not path_to_check.is_absolute():
            path_to_check = Path(basedir) / path_to_check
        if not path_to_check.exists():
            print(f'{typ} {path_to_check} existiert nicht.')
            sys.exit(1)
        return path_to_check


    def __init__(self):
        pass


@click.command()
def run():
    """Modulare Verwaltung von DNS-Zonen"""
    converter = FragenConv()
    with open('d:/TPSSE/github/Pruefungsfragen_TPSSE/Fragenpool/fragen_tpsse.yml', 'r') as f:
        data = yaml.load(f, Loader)
    print(data)

def main():
    # load_env()
    yaml.add_constructor('!include', construct_include, Loader)
    run()


if __name__ == '__main__':
    main()

