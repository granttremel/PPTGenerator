# -*- coding: utf-8 -*-
"""
Created on Wed Jan  3 10:47:59 2024

@author: gtremel
"""

import re
import os
from os import path

#%%

class cstringmanager:
    
    strlens = {
        'S':2,
        'C':4,
        'P':3,
        'N':3,
        'F':6,
        'Z':4
        }
    
    ptrns = {
        'date':'^([0-9]{8})_',
        'time':'_([0-9]{6})$|_',
        'S': '_S([0-9])_',
        'C': '_C([0-9]{3})_',
        'P': '_P([0-9]{2})_',
        'N': '_N([0-9]{2})_',
        'F': '_F([0-9]{5})_',
        'Z': '_Z([0-9]{3}).',
        }
    
    subdirs = ['Images','ProjDir_grant']
    
    imdir_format = os.path.join('{expdir}','{date}_{time}','{date}_{time}_S{S}','Images','{fname}')
    projdir_format = os.path.join('{expdir}','{date}_{time}','{date}_{time}_S{S}','ProjDir','{fname}')
    
    def __init__(self,**kwargs):
        
        self.dir = None
        self.predicate = lambda x:True
        
        #vals indexed by att index
        self.vals = list()
        for key in cosmx_code.keys:
            if key in kwargs:
                val = kwargs[key]
                if not hasattr(val,'__iter__'):
                    val = [val]
            else:
                val = [None]
            self.vals.append(val)
        
        # if 'W' in kwargs:
        #     self.vals.append(kwargs['W'])
        
        for otherkey in cosmx_code.otherkeys:
            if otherkey in kwargs:
                setattr(self,otherkey,kwargs[otherkey])
        
        if not 'W' in kwargs:
            self.useW = False
        else:
            self.useW = True
        
        self.lens = [len(val) for val in self.vals if val[0] is not None]
        self.ptrs = [0 for val in self.vals if val[0] is not None]
        
        #indexing ptrtx with att index gives ptr index
        self.ptrtx = list(range(len(self.ptrs)))
        #indexing invptrtx with ptr index gives att index
        self.invptrtx = [self.ptrtx.index(i) for i in range(len(self.ptrtx))] #attribute -> ptr
        
        self.groups = tuple()
        
        self.done = False
    
    def set_dir(self,_dir):
        
        ftot = str(_dir)
        
        self.validate_dir(ftot)
        
        pdir,last = os.path.split(ftot)
        for sd in self.subdirs:
            if sd == last:
                _dir = pdir
                
        
        self.dir = _dir
    
    def validate_dir(self, _dir):
        
        potentials = []
        
        for a,b,c in os.walk(_dir):
            # print(b)
            # for bdir in b:
            if 'CellStatsDir'in b:
                # print(b)
                anchor = a
                potentials.append(anchor)
        
        if len(potentials) == 0:
            raise Exception('No experiment directory identified. Verify the path or try a different one')
        
        expdirs = [os.path.dirname(os.path.dirname(anchor)) for anchor in potentials]
        expdirs= list(set(expdirs))
        
        if not len(expdirs) == 1:
            
            print(expdirs)
            raise Exception('huh')
        else:
            self.expdir = expdirs[0]
        
        timestamps = sorted(list(set([os.path.split(os.path.dirname(anchor))[1] for anchor in potentials])))
        
        dates = [re.search(self.ptrns['date'],ts).group(1) for ts in timestamps]
        times = [re.search(self.ptrns['time'],ts).group(1) for ts in timestamps]
        
        # print(self.expdir)
        # print(timestamps)
        
        self.date = dates[-1]
        self.time = times[-1]

    def set_predicate(self,pred_func):
        self.predicate = pred_func
        
    def set_date(self,date):
        self.date = date
    
    def set_time(self,time):
        self.time = time

    def __iter__(self):
        return self
    
    def __next__(self):
        if self.done:
            self.reset_pointers()
            raise StopIteration
        
        out = self.get_active()
        self.iterate_pointers()
        
        if not self.predicate(out):
            out = self.__next__()
        
        return out
    
    def iterate_until_valid(self):
        out = self.get_active()
        while not self.predicate(out):
            next(self)
            out = self.get_active()
        
            if self.done:
                return None
        
        return out
            
    def set_groups(self,*args):
        self.groups = list(args)
        nonargs = [k for k in cosmx_code.keys if k not in self.groups]
        
        if not self.useW:
            if 'W' in self.groups:
                self.groups.remove('W')
            else:
                nonargs.remove('W')
        
        attorder = nonargs + list(args)
        #indexing ptrtx with att index gives ptr index
        self.ptrtx = [cosmx_code.keys.index(att) for att in attorder]
        #indexing invptrtx with ptr index gives att index
        self.invptrtx = [self.ptrtx.index(i) for i in range(len(self.ptrtx))]
        self.lens = [len(self.vals[i]) for i in self.ptrtx]
    
    def estimate_ingroup_size(self):
        mylens = self.lens[len(self.vals)-len(self.groups):]
        out = 1
        for val in mylens:
            out*=val
        
        return out
    
    def estimate_outgroup_size(self):
        mylens = self.lens[:len(self.vals)-len(self.groups)]
        
        out = 1
        for val in mylens:
            out*=val
            
        return out
    
    def reset_pointers(self):
        self.ptrs = [0 for val in self.vals]
        self.done = False
        
    def reset(self):
        self.reset_pointers()
        
    def iterate_pointers(self):
        overflow = 0
        
        #logic for lsb
        self.ptrs[-1] += 1
        if self.ptrs[-1] == self.lens[-1]:
            self.ptrs[-1] = 0
            overflow = 1
        
        if overflow == 0:
            return
        
        #logic for remaining bits
        for i in range(-2,-len(self.vals)-1,-1):
            self.ptrs[i] += overflow
            
            if self.ptrs[i] == self.lens[i]:
                self.ptrs[i] = 0
                overflow = 1
            else:
                overflow = 0
                
            if overflow == 0:
                return
        
        self.done = True
        # raise StopIteration
        
        
    def get_active(self):
        inds = dict(zip(cosmx_code.keys, [val[self.ptrs[i]] for val,i in zip(self.vals,self.invptrtx)]))
        inds.update(dict(zip(cosmx_code.otherkeys, [getattr(self,k) for k in cosmx_code.otherkeys])))
        return cosmx_code(**inds)
    
    def get_all(self):
        outlist = [code for code in self]
        self.reset_pointers()
        return outlist
    
    @classmethod
    def get_dark(cls,code, prev = True,first = False):
        c = code.copy()
        if first:
            c.C = 0
            c.N = 0
        elif prev:
            if c.N %2 == 1:
                c.N -= 1
            else:
                pass
            
            if c.N == 0:
                c.C = 0
                
        else:
            if c.N %2 == 1:
                c.N += 1
            else:
                pass
        
        return c

    @classmethod
    def code_to_str(cls, code, datestr = None, timestr = None, proj = False, W = False):
        codestrs = [cls.q_str(char,q) for char,q in zip(cosmx_code.keys,code)]
        
        if proj:
            del codestrs[5]
        
        if not datestr is None and not timestr is None:
            # if not isinstance(datestr,str):
            datestr = str(datestr)
            timestr = str(timestr)
            codestrs = [datestr, timestr] + codestrs
        
        if not W:
            del codestrs[-1]
        
        return '_'.join(codestrs)

    @classmethod
    def search_for_code(cls,fdir,code):
        f = os.listdir(fdir)
        ptrn = '([0-9]{8})_([0-9]{6})_' + cls.code_to_str(code)
        matches = list()
        
        for file in f:
            # fdir,fname = path.split(file)
            res = re.search(ptrn,file)
            if not res is None:
                matches.append(path.join(fdir,file))
                
        return matches
    
    @classmethod
    def from_directory(cls,fdir):
        files = os.listdir(fdir)
        
        codes = list()
        
        date = None
        time = None
        
        for file in files:
            
            _,fname = os.path.split(file)
            filedict = {}
            
            for k,ptrn in cls.ptrns.items():
                res = re.search(ptrn,fname)
                if not res:
                    # continue
                    filedict[k] = None
                else:
                    filedict[k] = int(res.group(1))
            # print(filedict)
            ccode = cosmx_code(**filedict)
            codes.append(ccode)
            
            if not date and 'date' in filedict:
                date = filedict['date']
                
            if not time and 'time' in filedict:
                time = filedict['time']

        if not date:
            date = "0"*8
        
        if not time:
            time = "0"*6
        
        cman = cls.from_codelist(codes)
        cman.set_dir(fdir)
        cman.date = cls.zero_fill('',str(date),8 - len(str(date)))
        cman.time = cls.zero_fill('',str(time),6 - len(str(time)))
        
        return cman
    
    @classmethod
    def from_codelist(cls,codelist):
        spandict = {}
        for code in codelist:
            for key in code.keys:
                if not key in spandict:
                    spandict[key] = list()
                
                val = getattr(code,key)
                if (val is not None) and not val in spandict[key]:
                    spandict[key].append(val)
        
        cman = cls(**spandict)
        return cman
    
    def code_to_fname(self, code, proj = False, W = False):
        fname = type(self).code_to_str(code,datestr = self.date, timestr = self.time, proj = proj, W = W) + '.TIF'
        
        return fname
    
    def code_to_fpath(self, code, proj = False, W = False):
        
        fname = type(self).code_to_str(code,datestr = self.date, timestr = self.time, proj = proj) + '.TIF'
        
        if proj:
            fpath = self.projdir_format.format(expdir = self.expdir, date = self.date,time = self.time, S = code.S, fname = fname)
        else:
            fpath = self.imdir_format.format(expdir = self.expdir, date = self.date,time = self.time, S = code.S, fname = fname)
            
        return fpath
        

    @classmethod
    def zero_fill(cls,pre,suff,nzeros):
        return pre + "0"*nzeros + suff
    
    @classmethod
    def char_str(cls,char,n,outlen):
        nstr = str(n)
        nzeros = outlen - len(char) - len(nstr)
        # nzeros = outlen
        if nzeros < 0:
            raise Exception('string not long enough for characters')
        return cls.zero_fill(char,nstr,nzeros)
    
    @classmethod
    def fov_str(cls,f):
        return cls.char_str('F',f,cls.strlens['F'])
    
    
    @classmethod
    def z_str(cls,z):
        return cls.char_str('Z',z,cls.strlens['Z'])
    
    
    @classmethod
    def slot_str(cls,s):
        return cls.char_str('S',s,cls.strlens['S'])
    
    
    @classmethod
    def cycle_str(cls,c):
        return cls.char_str('C',c,cls.strlens['C'])
    
    @classmethod
    def pool_str(cls,p):
        return cls.char_str('P',p,cls.strlens['P'])
    
    
    @classmethod
    def rep_str(cls,n):
        return cls.char_str('N',n,cls.strlens['N'])
    
    @classmethod
    def chan_str(cls,w):
        return f'W{w}'
    
    @classmethod
    def q_str(cls,char,q):
        if char.upper() == 'F':
            return cls.fov_str(q)
        elif char.upper() == 'Z':
            return cls.z_str(q)
        elif char.upper() == 'S':
            return cls.slot_str(q)
        elif char.upper() == 'C':
            return cls.cycle_str(q)
        elif char.upper() == 'P':
            return cls.pool_str(q)
        elif char.upper() == 'N':
            return cls.rep_str(q)
        elif char.upper() == 'W':
            return cls.chan_str(q)
    
    
    @classmethod
    def build_filestr(cls,**kwargs):
        strlist = list()
        
        if 'time' in kwargs:
            t = kwargs['time']
        else:
            t = '00000000'
        
        if 'date' in kwargs:
            d = kwargs['date']
        else:
            d = '000000'
        
        strlist.append(d)
        strlist.append(t)
            
        for q in cls.strlens.keys():
            if q.upper() in kwargs or q.lower() in kwargs:
                qstr = cls.q_str(q,kwargs[q.upper()])
            else:
                qstr = cls.q_str(q,0)
            strlist.append(qstr)
            
        if 'fext' in kwargs:
            fext = kwargs['fext']
            # strlist.append(fext)
        else:
            fext = ''
    
        return '_'.join(strlist) + fext

    def __len__(self):
        l = 1
        for k in self.vals:
            l *= len(k)
            
        return l
    
    def __contains__(self, other):

        self.reset_pointers()
        
        val = False
        for code in self:
            val = val or (other == code)
        
        return val
            
        

class cosmx_code:
    
    keys = ('S','C','P','N','F','Z','W')
    otherkeys = ('date','time')
    
    def __init__(self,*args, **kwargs):
        
        if args and kwargs:
            raise ValueError('Only pass args or kwargs, not both')
        elif args:
            if not len(args) == len(self.keys):
                raise IndexError
            else:
                for key,val in zip(self.keys,args):
                    setattr(self,key,val)
        
        elif kwargs:
            for key in self.keys:
                setattr(self,key,kwargs.get(key,-1))
                    
            for otherkey in self.otherkeys:
                setattr(self,otherkey,kwargs.get(otherkey,None))
    
    @classmethod
    def from_tuple(cls,tup):
        if len(tup) == len(cls.keys):
            return cls(**dict(zip(cls.keys,tup)))
        elif len(tup) == len(cls.keys) + 1:
            return cls(**dict(zip(cls.keys + ('W',),tup)))
        else:
            raise IndexError
    @classmethod
    def from_args(cls,*args):
        if not len(args) == len(cls.keys):
            raise IndexError
        else:
            return cls(**dict(zip(cls.keys,args)))
    
    def to_dict(self):
        return {k:self[k] for k in self.keys}
    
    def to_tuple_except(self,*args):
        valid_keys = tuple(k for k in self.__dict__.keys() if not k in args)
        return tuple(getattr(self,k)for k in valid_keys)
    
    def to_tuple(self,*args):
        if not args:
            args = self.keys + self.otherkeys
        
        valid_keys = tuple(k for k in self.__dict__.keys() if k in args)
        return tuple(getattr(self,k) for k in valid_keys)
        
        
    @classmethod
    def test_equality(cls, cs1,cs2):
        t = True
        for i in range(len(cls.keys)):
            if cs1[i] < 0 or cs2[i] < 0:
                continue
            else:
                t = t and (cs1[i] == cs2[i])
        return t
            
    def copy(self):
        return type(self)(**self.__dict__)
    
    def __getitem__(self,key):
        if isinstance(key,str):
            return getattr(self,key)
        
        elif isinstance(key,int):
            return getattr(self,self.keys[key])

    def __str__(self):
        return ' '.join([f'{k}={self[k]}' for k in self.keys])
    
    def to_string(self,*args):
        if not args:
            args = self.keys
            
        args_ord = [k for k in self.keys if k in args]
        
        return ' '.join([f'{k}={self[k]}' for k in args_ord])
    
    def to_string_except(self,*args):
        
        if not args:
            args = self.keys
            
        args_ord = [k for k in self.keys if not k in args]
        
        return ' '.join([f'{k}={self[k]}' for k in args_ord])
    def __repr__(self):
        _classname = self.__class__.__name__
        mod = self.__module__
        # return _class
        loc = id(self)
        selfstr = self.__str__()
        
        return f'<{mod}.{_classname} {selfstr} at 0x{loc:x}>'
        # return f'<{mod}.{_classname} {selfstr} at 0x{loc:016x}>'
    def __hash__(self):
        return hash(tuple(self))
    
    def __eq__(self,other):
        
        if isinstance(other,tuple):
            other = type(self)(other)
        elif isinstance(other,self.__class__):
            pass
        else:
            return False
        
        return self.test_equality(self,other)
    
    @classmethod
    def intersection(cls,*codes):
        
        out = set()
        for k in cls.keys:
            vals = [codes[0][k] == code[k] for code in codes[1:]]
            res = all(vals)
            if res:
                out.add(k)
        return out
    
    @classmethod
    def symmetric_difference(cls,*codes):
        
        out = set()
        for k in cls.keys:
            vals = [not codes[0][k] == code[k] for code in codes[1:]]
            res = all(vals)
            if res:
                out.add(k)
        return out