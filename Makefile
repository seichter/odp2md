

build: upgrade-build
	python -m build


upgrade-build:
	python -m pip install --upgrade build

staging: clean
	python -m pip install .


clean:
	python -m pip uninstall odp2md
