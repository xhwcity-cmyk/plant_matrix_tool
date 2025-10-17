# setup.py
from setuptools import setup, find_packages

setup(
    name='SpeciesProcessor',
    version='1.0.0',
    description='物种数据整理工具',
    author='Your Name',
    author_email='your.email@example.com',
    packages=find_packages(),
    install_requires=[
        'openpyxl>=3.0.10',
    ],
    entry_points={
        'console_scripts': [
            'species-processor=species_processor:main',
        ],
    },
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)