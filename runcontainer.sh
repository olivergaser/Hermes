#!/bin/bash
docker build -t eml-converter .
docker run --rm -v $(pwd)/eml:/data eml-converter --input /data/