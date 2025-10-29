build: upgrade-build
	python -m build

venv:
	- deactivate
	rm -rf .venv
	python -m venv .venv

upgrade-build:
	python -m pip install --upgrade build

staging: clean build
	python -m pip install .

clean:
	yes | python -m pip uninstall odp2md
