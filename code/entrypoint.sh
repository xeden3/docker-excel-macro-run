#!/bin/bash

# 将脚本参数传递给 xvfb-run、wine 和 python 程序
# disable the 'X connection to :100 broken (explicit kill or server shutdown).' warning 
xvfb-run -a wine python /opt/wineprefix/drive_c/app/excel_xlsm_macro_run.py "$@" | grep -v '100 broken'
