from distutils.core import setup
import py2exe
import numpy

setup(
    options={'py2exe': {'bundle_files': 1,
                        'compressed': True,
                        "includes": ["sip"]
                        }},
    console=["telefony.py"],
    zipfile=None,
)
