from setuptools import setup

install_requires = [
    "openpyxl==3.0.5"
]

setup(
    name="openpyxl_custom_use",
    version="0.0.1",
    url="https://github.com/gyungsumin/custompackage-use-openpyxl-easily.git",
    author="gyungsumin",
    author_email="gyungsumin@gmail.com",
    license="gyungsumin",
    packages=["openpyxl_custom_use"],
    install_requires=install_requires
)
