#!/usr/bin/python
# coding: utf-8
# By Sébastien Brière - 2016
# Importations de documents PDF

import lib.importalf

try:
    instance = lib.importalf.importAlf()
    
    instance.gen_dialog()
    
except Exception, e:
    print str(e)


