"""
** USER STORY ***********************

I am a scientist at nanostring. I just completed a new cosmx 2.0 experiment and
I want to verify the following
    a looks good
    b looks great
    compare c with d
    I don't care about e
    I was worried about f



"""


#%%

from glib import ppt
from glib import cosmx_strings

from glib import gtools as gt

#%%

exppath = r"\\cosmx-0199\F\RunNC78_20240322_2309h0199_NC78-CPA-6colorexpt3"

S = [1,2,3,4]
C = 1
P = 0
N = [1,17,33,43,53]
F = range(1,8)
Z = 6
W = ['BB','GG','YY','RR','RN','BG']

cstr = cosmx_strings.cstringmanager(S=S,C=C,P=P,N=N,F=F,Z=Z,W=W)

cstr.set_dir(exppath)

cstr.set_groups('N','F','W')

ppt_output_path = gt.fp(r'C:\\Dash\\Data\\example1.pptx')
template_path = r'C:\Users\gtremel\Documents\Python Scripts\pptx_nstg_formats'

#%%

params = ppt.pptgeneratorparams(
    margins = [1.0,0.0,0.0,0.0],
    sep = 0.0,
    crop = [(0.4,0.6),(0.4,0.6)],
    tile_cols = 6,
    auto_labels = True,
    max_rows = 4,
    contrast_broadcast_mode = -1,
    contrast_group = ('S','W'),
    contrast_calculate_mode = 3,
    max_per_slide = -1,
    dark_mode = True,
    save_pngs = True
    )

contrast_data = [
    (0,99.5), #BB
    (1,99), #GG
    (0,97),  #YY
    (0,99), #RR
    (2,99), #RN
    (1,99.5) #BG
    ]

#%%
pptg = ppt.pptgenerator(ppt_output_path, cstr, params, contrast_data)
pptg.set_template_path(template_path)

#%%
# pptg.set_contrast_data(contrast_data)
# pptg.generate_first()
pptg.generate(1)
