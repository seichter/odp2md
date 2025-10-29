

build: upgrade-build
	python -m build


upgrade-build:
	python -m pip install --upgrade build

staging: clean build
	python -m pip install .


clean:
	yes | python -m pip uninstall odp2md
