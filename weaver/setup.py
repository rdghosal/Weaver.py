from distutils.core import setup

setup(
    name="weaver",
    version="0.1",
    package_dir={
        "weaver": "", 
    },
    packages=["weaver", "weaver.reports", "weaver.reports.sim", "weaver.reports.conf"]
)