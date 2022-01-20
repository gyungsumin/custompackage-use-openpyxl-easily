from setuptools import setup

install_requires = [
    "openpyxl==3.0.5"
]

setup(
    name="excel",
    version="0.0.1",
    url="https://github.com/gyungsumin/excel.git",
    author="gyungsumin",
    author_email="gyungsumin@gmail.com",
    license="gyungsumin",
    packages=["excel"],
    install_requires=install_requires
)
