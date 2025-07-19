from setuptools import setup, find_packages

setup(
    name="salesforce_data_migrator",
    version="0.1.0",
    description="GUI tool per migrazione record Salesforce",
    author="simongionni",
    author_email="simonjohnny99@gmail.com",
    packages=find_packages(),        # trova record_migrator
    include_package_data=True,
    install_requires=[
        "pandas",
        "openpyxl",
        "simple-salesforce",
        "PyYAML",
        "ttkthemes",
    ],
    entry_points={
        "console_scripts": [
            "salesforce-data-migrator = record_migrator.main:main",
        ],
    },

)
