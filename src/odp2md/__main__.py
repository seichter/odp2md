import os, sys


if not __package__:
    # Make CLI runnable from source tree with
    #    python src/package
    package_source_path = os.path.dirname(os.path.dirname(__file__))
    sys.path.insert(0, package_source_path)


def main():
    from .odp2md import main_cli

    main_cli()


if __name__ == "__main__":
    main()
