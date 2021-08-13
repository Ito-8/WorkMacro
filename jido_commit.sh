#!/bin/bash

rm -rf ./src/WorkMacro.xlsm

cscript vbac.wsf decombine

git add .
git commit -m $1
git push
