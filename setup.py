import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(name='PyWrapOrigin',
        version='1.0.0',
        description='A python wrapper that simplifies sending data to and plotting in OriginLab from python, and it allows plotting OriginLab graphs without needing a graph template.',
        long_description=long_description,
        url='https://github.com/chrislauyc/PyWrapOrigin.git',
        author='Chris Lau',
        author_email='chrislyc.lau@gmail.com',
        packages=setuptools.find_packages(),
        py_modules=['PyWrapOrigin'],
        include_package_data=True,
        package_data={
            '':['DataPlotter.h','DataPlotter.cpp','Plotter.cpp'],
            },
        #scripts=['PyWrapOrigin/PyWrapOrigin.py'],
        classifiers=[
            'Programming Language :: Python :: 3',
            'License :: OSI Approved :: MIT License',
            'Operating System :: Microsoft :: Windows'
            ]
        )

#have to run ```pip install .```
