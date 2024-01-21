from os import path
from setuptools import setup, find_packages

"""
    Setup file
    License: GNU
"""

# Content of the README file
current_path = path.abspath(path.dirname(__file__))
with open(path.join(current_path, "README.md")) as f:
    long_description = f.read()


def read_requirements_file(path):
    requires = list()
    f = open(path, "rb")
    for line in f.read().decode("utf-8").split("\n"):
        line = line.strip()
        if line:
            requires.append(line)
    return requires

REQUIRES = read_requirements_file("requirements.txt")

# Setup
setup(
	name='Ideal-Home-Location-Matcher',
	version='1.0',
	author='Andrew Drogalis',
    author_email='andrew.drogalis2@gmail.com',
    description="A collection of weather, natural disaster, and US Census data processed and ranked to find the ideal home location for each individual's preferences.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/andrew-drogalis/Ideal-Home-Location-Matcher',
    python_requires=">=3.6.0",
    packages=find_packages(),
    include_package_data=True,
    install_requires=REQUIRES,
    license="GNU"
)
