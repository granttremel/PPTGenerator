# -*- coding: utf-8 -*-
"""
User Story

"""

from PPTGenerator.util import fp
from PPTGenerator import ppt
from PPTGenerator import cosmx_strings

#%%
exppath = r"\\cosmx-0199\F\RunNC78_20240322_2309h0199_NC78-CPA-6colorexpt3"

S = [1,3,4]
C = [1,2]
P = 0
N = range(1,51,4)
F = [1,2,3,4,5,6,7]
Z = 6
W = ['GG','YY']

cstr = cosmx_strings.cstringmanager(S=S,C=C,P=P,N=N,F=F,Z=Z,W=W)

cstr.set_dir(exppath)
cstr.set_groups('N','F')

#%%

pptfp = fp(r'C:\Dash\Data\example4.pptx')
template_path = r'C:\Users\gtremel\Documents\Python Scripts\pptx_nstg_formats'
params = ppt.pptgeneratorparams(
    margins = [1.0,0.0,0.0,0.0],
    crop = [(0.45,0.55),(0.4,0.6)],
    tile_cols = 7,
    auto_labels = True,
    sep = 0,
    contrast_group = ('F','W'),
    contrast_broadcast_mode = -1,
    contrast_calculate_mode = 0,
    max_rows = 7,
    dark_mode = True,
    save_pngs = False
    )


contrast_data = [
    (1,99), 
    ]

pptg = ppt.pptgenerator(pptfp,cstr,params, contrast_data = contrast_data)
pptg.set_template_path(template_path)


#%%

pptg.generate()
