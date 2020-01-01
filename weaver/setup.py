from setuptools import setup

setup(
    name="weaver",
    version="0.1",
    package_dir={
        "weaver": "",
        "weaver.reports": "reports",
        "weaver.reports.sim": "reports/sim" 
    },
    packages=["weaver", "weaver.reports", "weaver.reports.sim"],
    entry_points={
        "console_scripts": ["weaver=app:main"]
    }
)