from setuptools import setup, find_packages

setup(
    name="broadcast",
    version="0.1.0",
    description="A Python package for Broadcast Excel Addin automation using xlwings",
    author="Luan Correia",
    author_email="luan.a.correialive@gmail.com",
    url="https://github.com/CorreiaLuan/broadcast",
    packages=find_packages(),
    install_requires=["xlwings", "pandas"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
)
