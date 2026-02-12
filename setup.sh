#!/bin/bash
apt-get update
apt-get install -y libxrender1 libxext6 libxcb1 libx11-6 libxau6 libxdmcp6 libxcb-render0 libxcb-shm0 libfreetype6 libfontconfig1
pip install --upgrade pip
pip install -r requirements.txt
