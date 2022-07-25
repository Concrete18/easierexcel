from setuptools import setup, find_packages

classifiers = [
    "Development Status :: 5 - Production/Stable",
    "Intended Audience :: Developers",
    "Operating System :: Microsoft :: Windows :: Windows 10",
    "License :: OSI Approved :: MIT License",
    "Programming Language :: Python :: 3",
]

setup(
    name="easierexcel",
    version="0.9.2",
    description="Easy viewing/editing of Excel Files",
    long_description=open("README.md").read(),
    # long_description=open('README.md').read() + '\n\n' + open('CHANGELOG.txt').read(),
    long_description_content_type="text/markdown",
    url="https://github.com/Concrete18/easierexcel",
    author="Concrete18",
    author_email="<michaelericson19@gmail.com>",
    license="MIT",
    classifiers=classifiers,
    keywords="excel",
    packages=find_packages(),
    install_requires=["openpyxl", "pandas"],
)
