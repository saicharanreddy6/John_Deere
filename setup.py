from setuptools import setup, find_packages

setup(
    name="xml_processor",
    version="0.1",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=["pandas"],
    entry_points={
        "console_scripts": [
            "xml_processor=xml_processor:main",
        ],
    },
)
