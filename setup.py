import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name = 'odp2md',
    version = '0.3.2',
    description = "convert OpenOffice odp to markdown",
    long_description = long_description,
    long_description_content_type="text/markdown",
    url = "https://github.com/seichter/odp2md",
    author = "Hartmut Seichter",
    author_email = "hartmut@technotecture.com",
    packages = setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Operating System :: OS Independent"
    ],
    python_requires='>=3.6',
    scripts=['bin/odp2md']
)
