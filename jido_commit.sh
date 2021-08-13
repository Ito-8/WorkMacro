#!/bin/bash

cscript vbac.wsf decombine

git add -A .
git commit -m < $1
git push
