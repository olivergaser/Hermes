#!/bin/bash

rm /eml/*.pdf
rm /eml/*.tif
./venv/bin/python -m unittest tests/test_conversion.py