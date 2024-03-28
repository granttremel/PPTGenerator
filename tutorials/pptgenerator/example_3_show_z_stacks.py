# -*- coding: utf-8 -*-
"""
User Story

"""

from glib import gtools as gt
from glib import ppt
from glib import cosmx_strings

#%%
exppath = r"\\cosmx-0199\F\RunNC78_20240322_2309h0199_NC78-CPA-6colorexpt3"

S = [4]
C = [2]
P = 0
N = [27]
F = [1,2,3,4,5,6,7]
Z = range(1,9)
W = ['BB','GG','YY','RR','RN','BG']

cstr = cosmx_strings.cstringmanager(S=S,C=C,P=P,N=N,F=F,Z=Z,W=W)

cstr.set_dir(exppath)
cstr.set_groups('Z')

#%%

pptfp = gt.fp(r'C:\\Dash\\Data\\example3.pptx')
template_path = r'C:\Users\gtremel\Documents\Python Scripts\pptx_nstg_formats'
params = ppt.pptgeneratorparams(
    margins = [1.0,0.1,0.1,0.1],
    crop = [(0.4,0.6),(0.4,0.6)],
    tile_cols = 4,
    auto_labels = False,
    sep = 0,
    contrast_group = ('F','W'),
    contrast_broadcast_mode = -1,
    contrast_calculate_mode = 0,
    max_rows = 2,
    dark_mode = True,
    save_pngs = False
    )


contrast_data = [
    (1,99.5), #BB
    (1,97), #GG
    (1,97), #YY
    (1,99), #RR
    (0,99), #RN
    (1,99.5) #BG
    ]

pptg = ppt.pptgenerator(pptfp,cstr,params, contrast_data = contrast_data)
pptg.set_template_path(template_path)

pptg.set_labels(rowlabels = ['Z=1-4','Z=5-8'])

#%%

pptg.generate(1)
