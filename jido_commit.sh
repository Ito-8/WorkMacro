#!/bin/bash

rm -rf ./src/WorkMacro.xlsm

cscript vbac.wsf decombine

git add -A .
git commit -m $1
# git push
