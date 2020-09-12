import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(name='PyWrapOrigin',
        version='0.0.1',
        description='A python wrapper that simplifies sending data to and plotting in OriginLab from python, and it allows plotting OriginLab graphs without needing a graph template.',
        long_description=long_description,
        url='https://github.com/chrislauyc/PyWrapOrigin.git',
        author='Chris Lau',
        author_email='chrislyc.lau@gmail.com',
        packages=setuptools.find_packages(),
        classifiers=[
            'Programming Language :: Python :: 3',
            'License :: OSI Approved :: MIT License',
            'Operating System :: Microsoft :: Windows'
            ]
        )