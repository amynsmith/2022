import re
import re
import subprocess
import math
import numpy as np
import os

import logging
import logging.config
import yaml
# logging.basicConfig(level=logging.ERROR)
with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
# logger = logging.getLogger(__name__)
logger = logging.getLogger('verboseLogger')
# logger.setLevel(logging.DEBUG)


def writePagenumList(toConvertPages, toconvertfn):
    toConvertPages.sort()
    logger.info(f'toConvertPages list size: {len(toConvertPages)}')
    # Command too long: -sPageList=
    nchunks = math.ceil(len(toConvertPages) / 20)
    logger.info(f'divided into {nchunks} chunks')
    lists_toConvertPages = np.array_split(toConvertPages, nchunks)
    pagenumlist = []
    for sublist in lists_toConvertPages:
        tmpstr = ",".join([str(i) for i in sublist])
        pagenumlist = pagenumlist + [tmpstr]
    with open(toconvertfn, "w", encoding="utf-8") as fp:
        content = "\n".join(pagenumlist)
        fp.writelines(content)
    return pagenumlist


def loadPagenumList(toconvertfn):
    txt = open(toconvertfn, "r", encoding="utf-8")
    line = txt.readline()
    pagenumlist = []
    while line:
        tmpstr = line.strip()
        if tmpstr:
            pagenumlist.append(tmpstr)
        line = txt.readline()
    txt.close()
    return pagenumlist


def loadPagenums(toconvertfn):
    pagenumlist = loadPagenumList(toconvertfn)
    tmpstr = ",".join(pagenumlist)
    tmplist = tmpstr.split(",")
    pagenums = [int(i) for i in tmplist]
    return pagenums
