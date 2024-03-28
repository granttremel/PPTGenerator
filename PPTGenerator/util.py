#%%imports
import matplotlib as mpl
import matplotlib.pyplot as plt 

import numpy as np
import os
import os.path as pth
import cv2
from PIL import Image
from PIL.TiffTags import TAGS
import tifffile as tf
import imageio
import json

#%% classes

class fp:
    def __init__(self, ftot, ftag=''):
        d,f = pth.splitdrive(ftot)
        self.drive = d
        self._ftot = ftot
        self.ftag = ftag
        fpathname,fext = pth.splitext(f)

        self.fext = fext
        fpath,fname = pth.split(fpathname)
        self.fpath = fpath
        self.fname = fname
        self.fnx = self.fname + self.fext

        if not pth.exists(pth.join(self.drive,self.fpath)):
            os.mkdir(pth.join(self.drive,self.fpath))

    def insert_suffix(self, suffix):
        # if not self.isfile:
        #     return ''
        return pth.join(self.fpath, self.fname + '_' + suffix + self.fext)

    # def add_dir(self, dirname):
    #     return fd(pth.join(self.drive,self.fpath, dirname), self.fname)
    
    @property
    def ftot(self):
        return pth.join(self.drive,self.fpath, self.fname + self.fext)

    @classmethod
    def from_parts(cls, fpath, fname, fext, drive = ''):
        return cls(pth.join(drive,fpath, fname + fext))
    
    # @classmethod
    # def select(cls, start = None):
    #     Tk().withdraw()
    #     fname = askopenfilename()
        
    #     return cls(fname)

    def clone(self):
        return type(self)(self.ftot)

    def open_dir(self):
        os.startfile(pth.join(self.drive,self.fpath))

    def make(self):
        # fpwithdrive = os.path.join(self.drive,self.fpath)
        # print(fpwithdrive)
        if not os.path.exists(self.ftot):
            os.mkdir(self.ftot)

    def __str__(self):
        return self.ftot

    def __repr__(self):
        outstr = "{0}\n{1}".format(type(self), self.ftot)
        return outstr
    
    
#%% from improc

def scale_and_convert(im,vmin,vmax,newtype):
    try:
        typeinfo = np.iinfo(newtype)
    except Exception as e:
        raise e
    
    minval = typeinfo.min
    maxval = typeinfo.max
    
    outim = ((im.astype('double') - vmin) * (maxval - minval) / (vmax - vmin) + minval).clip(minval,maxval).astype('uint8')

    return outim

def scale_and_convert_8b(im,vmin,vmax):
    return scale_and_convert(im,vmin,vmax,np.uint8)

def crop_bds(im,bounds, relative = False):
    arrsz = im.shape
    
    for bd in bounds:
        if bd[0]-bd[1] <= 1:
            relative = True
    
    if len(arrsz) > 2:
        cropbds = (slice(None),) * (len(arrsz) - 2)
        imsz = arrsz[-2:]
    else:
        cropbds = tuple()
        imsz = arrsz

    if relative:
        imbds = tuple([slice(int(_min * sz),int(_max * sz)) for (_min,_max),sz in zip(bounds,imsz)])
        cropbds += imbds
    else:
        cropbds += tuple([slice(int(_min),int(_max)) for (_min,_max) in bounds])
 
    cropbds = tuple(reversed(cropbds))   
 
    return im[cropbds]

def read_im(fp):
    md = get_tiff_info(fp)
    if 'PageNumber' in md.keys():
        s,ims =read_stack(fp) 
        return s,ims
    else:
        return True,cv2.imread(str(fp), cv2.IMREAD_ANYDEPTH)
    
def write_im(fp,im, metadict = {}):
    flags = ()
    f = str(fp)
    return cv2.imwrite(f,im, flags)

def read_stack(fp,force_stack = True, inds = tuple()):
    md = get_tiff_info(fp)
    if 'PageNumber' in md.keys() or force_stack:
        s,ims = cv2.imreadmulti(str(fp),[],cv2.IMREAD_ANYDEPTH)
        # print(len(ims))
    else:
        print('Image detected to be single page')
        s,ims = read_im(fp)
    
    if inds:
        # print('inds provided')
        
        if len(ims) < len(inds):
            print('Too many indices provided...')
        elif hasattr(inds,'__iter__'):
            ims = [ims[ind] for ind in inds]
        else:
            ims = ims[inds]
            
        if len(ims) == 1:
            ims = ims[0]
    
    return s,np.array(ims)

def read_multichan_substack(fp, keys = []):
    
    outdict = dict()
    
    with tf.TiffFile(str(fp)) as im:
        npages = len(im.pages)
        md = im.pages[0].tags
        imdesc = json.loads(md['ImageDescription'].value)
        
        if 'ChannelOrder_ExEm' in imdesc:
            labels = imdesc['ChannelOrder_ExEm']
        elif 'ChannelOrder' in imdesc:
            labels = [a+a for a in imdesc['ChannelOrder']]
            
        if not keys:
            keys = labels
            
        for key in keys:
            if not key in labels:
                print(f'Key {key} not located in image. Skipping...')
                continue
            
            ind = labels.index(key)
            pagedata = im.pages[ind].asarray()
            outdict[key] = pagedata
            
    return outdict

def write_png(_fp,im,vmin, vmax):
    
    if not isinstance(_fp,fp):
        _fp = fp(_fp)
    
    _fp.fext = '.png'
    
    im_8b = scale_and_convert_8b(im,vmin,vmax)
    
    with imageio.get_writer(fp.ftot,mode = 'I') as writer:
        writer.append_data(im_8b)



def get_tiff_info(fp):
    with Image.open(str(fp)) as im:
        meta_dict = {TAGS[key] : im.tag[key] for key in iter(im.tag.keys())}
    return meta_dict


def get_image_description(fp):
    md = get_tiff_info(fp)
    if 'ImageDescription' in md.keys():
        return json.loads(md['ImageDescription'][0])
    else:
        raise KeyError('ImageDescription not present in tiff metadata')
        
def get_exems(fp):
    info = get_image_description(fp)
    
    if 'ChannelOrder_ExEm' in info:
        exems = info['ChannelOrder_ExEm']
    
    elif 'ChannelOrder' in info:
        exs = info['ChannelOrder']
        ems = exs
        exems = [ex + em for ex,em in zip(exs,ems)]
    else:
        exems = []
    
    return exems
        

#%% from plotting

def imshow(im,**kwargs):
    vrange = kwargs.get('vrange',(None,None))
        
    cmap = kwargs.get('cmap','gray')
    
    wc = kwargs.get('write_contrast',False)
    
    ch = kwargs.get('title',None)
    
    pct = kwargs.get('percentiles',None)
    
    fc = kwargs.get('facecolor', (0,0,0))
    
    show = kwargs.get('show', True)
    
    if not show:
        plt.ioff()
    
    # cbds = kwargs.get('crop_bounds',None)
    
    # if cbds:
    #     im = 

    if pct:
        vmin,vmax = np.percentile(im,pct)
        vmin,vmax = int(vmin),int(vmax)

    f,ax = plt.subplots(layout = 'tight')  
    f.set_facecolor(fc)      
    ax.imshow(im,vmin = vrange[0], vmax = vrange[1], cmap = cmap)
    ax.set_axis_off()
    f.subplots_adjust(bottom = 0, top = 1, left = 0, right = 1,hspace = 0, wspace = 0)
    # ax.xaxis.set_major_locator(plt.NullLocator())
    # ax.yaxis.set_major_locator(plt.NullLocator())
    
    if wc:
        tsz = 10
        # delta = 2*tz
        
        tmin = ax.text(0.99,0.05,
                       f'min = {vrange[0]:0.0f}',
                       ha = 'right', 
                       transform = ax.transAxes,
                       size = tsz,
                       bbox = dict(facecolor = 'white',alpha = 1, linewidth = 0, pad = 0.5))
        tmax = ax.text(0.99,0.01,
                       f'max = {vrange[1]:0.0f}',
                       ha = 'right',
                       transform = ax.transAxes,
                        size = tsz,
                       bbox = dict(facecolor = 'white',alpha = 1, linewidth = 0, pad = 0.5))
    
    if ch:
        tsz = 12
        
        tchan = ax.text(0.01,0.99,
                       ch,
                       ha = 'left',
                       va = 'top',
                       transform = ax.transAxes,
                       size = tsz,
                       bbox = dict(facecolor = 'white',alpha = 1, linewidth = 0, pad = 0.5))
    
        
    if not show:
        plt.ion()
        
    return f