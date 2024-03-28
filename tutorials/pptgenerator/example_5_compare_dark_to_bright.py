# -*- coding: utf-8 -*-
"""
User Story

"""

from PPTGenerator.util import fp
from PPTGenerator import ppt
from PPTGenerator import cosmx_strings


#%%
exppath = r"\\beta04\E\RunNC91_20240320_beta04_D-DNA-BlockingOligos-Exp2"

S = [1]
C = [1,2]
P = 0
N = range(12)
F = [1,6,7]
Z = 6
W = ['BB','GG','YY','RR']

cstr = cosmx_strings.cstringmanager(S=S,C=C,P=P,N=N,F=F,Z=Z,W=W)

cstr.set_dir(exppath)
cstr.set_groups('N','W')

#%%

pptfp = fp(r'C:\Dash\Data\example5.pptx')
template_path = r'C:\Users\gtremel\Documents\Python Scripts\pptx_nstg_formats'
params = ppt.pptgeneratorparams(
    margins = [1.0,0.0,0.0,0.0],
    crop = [(0.4,0.6),(0.4,0.6)],
    tile_cols = len(W),
    auto_labels = True,
    sep = 0,
    contrast_group = ('F','W'),
    contrast_broadcast_mode = -1,
    contrast_calculate_mode = 3,
    max_rows = 2,
    dark_mode = True,
    save_pngs = True
    )


contrast_data = [
    (1,99.5),
    (1,99),
    (1,99),
    (1,99.5)
    ]

pptg = ppt.pptgenerator(pptfp,cstr,params, contrast_data = contrast_data)
pptg.set_template_path(template_path)


#%%

pptg.generate()
