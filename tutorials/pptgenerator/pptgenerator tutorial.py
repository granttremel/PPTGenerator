"""
**********************************************
PPTGenerator tutorial
**********************************************

Step 0: Clone the PPTGenerator repository
1. Download git bash from https://git-scm.com/downloads
2. Navigate to the desired directory and input the command:
        git clone https://github.com/granttremel/PPTGenerator.git
3. Input credentials as necessary
4. Give my repos a star :)
    
5. If you are using an anaconda virtual environment, first create and activate the
    environment with python 3.11 as follows:
        conda create --name MyEnvName python=3.11
        conda activate MyEnvName
    Then, run the following commands:
        conda install pip
        pip install -r "<path-to-repos>/requirements.txt"
        conda develop "<path-to-repos>"

6. In the anaconda console with the venv activated, type python (or otherwise
    enter your python environment) and type:
        import PPTGenerator
        from PPTGenerator import ppt
    to verify the install. you shouldnt have to install any dependencies..
7. you may have to add the PPTGenerator directory to your PYTHONPATH
    (in spyder, Tools -> PYTHONPATH manager -> Add path)

"""

#%%
#these are the necessary imports
from PPTGenerator import ppt
from PPTGenerator import cosmx_strings

#not necessary but nice to have
from PPTGenerator import util

#%% cosmx string manager setup

#path to top level experiment directory
exppath = r"\\cosmx-0199\F\RunNC78_20240322_2309h0199_NC78-CPA-6colorexpt3"

"""
The following is how to define the cross-section of the run you would like displayed
Each variable corresponds to a parameter that is iterated during run progression
The variable should be defined as a single item or a collection of the valid 
values attained by the parameter. For example, Flow cells (S) should be 1-4,
channels should correspond to a valid channel combo, etc. Anything invalid will
be skipped
"""
#Flow cell
S = [2,3]
#Cycle
C = [1]
#Pool
P = 0
#Reporter
N = range(1,55,2)
#FOV
F = range(1,8)
#Z-slice
Z = []
#Channel
W = ['BB','GG','YY','RR','RN','BG']

#Initialize the cosmx string manager by assigning variables defined above to
#The SCPNFZW keyword argments. Do not skip any!
cstr = cosmx_strings.cstringmanager(S=S,C=C,P=P,N=N,Z=Z,F=F,W=W)

#This sets the directory of the cosmx string manager to the experiment directory
cstr.set_dir(exppath)

#%% Setting display groups
"""
The display groups determine which values are arranged into a grid within a slide.
All images which DIFFER at these values will be presented in the same slide. All 
images with these values will in common will be presented in different slides. 
This will also determine how many images will be shown in a slide, and how many
slides will be generated. For the example here, the images per slide will be 
len(F)*len(W) = 42, and the number of slides will be len(S)*len(C)*len(N).. = 189.
As a rule of thumb, you probably want to limit the display group size to be ~100,
or else the images will be very very small! Note that the ORDER MATTERS. 
The last argument provided (on the very right) will be iterated over first, 
then second from the right, etc. More on this later.
"""

#This will set the display groups to be by FOV and channel. This means that each
#slide will contain pictures across all FOV's and channels that were acquired.
#Because of the order that they are provided, rows will contain different channels
# and columns will contain different FOV's
cstr.set_groups('F','W')

#In this example, the groups are the same as above but the display order is
#different. Here, rows will contain different FOVs and columns will contain
#different channels
cstr.set_groups('W','F')

#This example will set the display groups to be reporter and Z-slice. So, every 
#FOV and channel will get their own slide, and in each slide the entire z-stack
# will be shown for all reporter pools.
cstr.set_groups('N','Z')

#Here the display groups are set to FOV, cycle, and reporter. The information
#contained in rows depends on the gridsize (described later) but this shows off
#the flexibility of the tool.
cstr.set_groups('F','C','N')

#let's stick with the easy example...
cstr.set_groups('F','W')
#%% Initializing the ppt generator

#these paths point to the ppt output and the templates, respectively. The ppt 
#output path should be local, with pptx extension. The necessary templates
#are available from my empshare:
#\\nano-data2\Emp_Share\GTremel\NSTG empty ppt formats
#please make local copies of these files! because of limitations of the 
#pptx library, these templates exactly are required to generate the output.
#ppt output path may be a gt.fp object or a raw file path.
ppt_output_path = util.fp(r'C:\\Dash\\Data\\nc78.pptx')
template_path = r'C:\Users\gtremel\Documents\Python Scripts\pptx_nstg_formats'

#%% The pptgenerator params
"""
The pptgenerator params comprise most of the customizability of the pptgenerator.
Each parameter has a specific function that is described in detail below. The 
defaults are usually a good starting point, so when in doubt, use them
"""
params = ppt.pptgeneratorparams.default()
for k,v in params.__dict__.items():
    print(f'{k}: {v}')
    
#These are the margins of the ppt slides that are generated. Each value corresponds
#to the margin size in inches, in the order top, left, bottom, right. The top
#margin of the NSTG template is about 1.1 inches, other values can be reasonably
#set to within 0.0 and 2.0.
params.margins = [1.0,0.5,0.5,0.5]

#This is the separator or the margin between adjacent images. Setting it to a
#small value can make a more attractive slide, but setting to 0 is more space
#efficient. Note that this is more like the minimum value of the separation
#between images, the actual separation may be larger if space allows.
params.sep = 0.0

#The crop parameter allows you to crop the images to a reduced size, showing 
#less of the FOV but increasing the size of the features to be visualized. The 
#values are in normalize image units, i.e. between 0 and 1. each tuple contains
#the minimum and maximum bound, and the two tuples correspond to the X and Y
#axes of the images, respectively. The example below means that the horizontal (X)
#direction will be cropped from the 20% point to 70%, so the final image will
#span 50% of the image width. Similarly, the vertical direction will be cropped
#from 30% to 60% and will span 30% of the image height. if set to [(0,1),(0,1)],
#the images will not be cropped.
#To display features on the scale of a few cells, the span should be ~10% in 
#both directions. and a 2:1 or 4:3 aspect ratio generally looks best
params.crop = [(0.2,0.7),(0.3,0.6)]

#the tile_cols parameter indicates how many columns the grid of images within a 
#slide will contain. The rows is then determined by the total number of images
#per slide. This number should usually correspond to the fastest iterating 
#argument provided to the display groups, e.g. above the groups were set to
# ('F','W'), so channels (W) are the fastest iterating group. Since there are 
# 6 channels in play (len(W) = 6), the best option for tile_cols is 6. Different
#values can give you access to greater variety of customization
params.tile_cols = len(W)

#This parameter tells the pptgenerator whether or not to attempt to generate
#auto labels. This works best when the rows and columns of the grid of images
#each correspond to one parameter, like with the current example, columns 
#are channels and rows are FOVs, so they will be labeled as such. Custom labels
#may also be set for more flexibility (later)
params.auto_labels = True

#This parameter tells the pptgenerator how to assign contrasts to images, and
#which types of images will recieve the same contrast settings (think imageJ
#"propagate to all other open images" but more flexible). There exist the following
#presets:
    # SELF = 0 (all images get their own contrasts)
    # BY_CHAN = 1 (contrasts are the same between channels)
    # BY_FOV = 2 (contrasts are the same between FOVs)
#to allow custom contrast broadcast groups, set to -1
params.contrast_broadcast_mode = -1

#This parameter is only relevant if the contrast broadcast mode is set to -1. 
#images that are common at all these values will have the same contrast settings,
#and images that differ at these values will have different contrast settings.
#in this example, images with the same FOV, reporter pool, and FC will have the 
#same contrast settings applied, but between channels, cycles and Z, will have 
#different contrast settings (but will be consistent within those other groups)
params.contrast_group = ('F','N','S')

#This parameter tells the pptgenerator how to calculate the contrast limits that
#are applied within each group. It assumes the following values:
    # FIRST = 0  just use the first image loaded (fast!)
    # MINMIN_MAXMAX = 1 maximum of maxima, minimum of minima
    # MAXMIN_MINMAX = 2 maximum of minima, minimum of maxima
    # AVG = 3 average across images in broadcast category
#I recommend 0 for fast ppt generation, but the other values can give you a lot of 
#customization
params.contrast_calculate_mode = 0

#These two parameters control how many rows and how many images are displayed
#on each slide. if the number of images decided by the display groups exceeds
#either of these values, then the grid will overflow to subsequence slides until
#the display group is completed. To disable each of these settings, set to -1,
#and only one of these values should be set! 
params.max_rows = 4
params.max_per_slide = -1

#This parameter will set the ppt to a dark style. this is easier on the eyes and
#enhances detail in the images. The title slide however will stay light.
params.dark_mode = False

#This parameter will decide whether or not to maintain a cache of png files when
#generating graphics. This option can make things faster and reduce memery
#usage when generating ppts with many images. However, with smaller ppts it will
#slow things down.
params.save_pngs = False

#the params can be assembled quickly as follows:
params = ppt.pptgeneratorparams(
    margins = [1.0,0.5,0.5,0.5],
    sep = 0.1,
    crop = [(0.2,0.7),(0.3,0.6)],
    tile_cols = 6,
    dark_mode = True
    )

#any keyword arguments provided will be used, and any not provided will fall 
#back to the defaults.

#params can be adjusted by directly setting attributes:
params.save_pngs = True


#%% Defining the contrast data
"""
When initializing the pptgenerator, it requires a contrast_data parameter with 
contrast bounds in units of percent. For context, contrast in images is usually 
defined as histogram percentiles. In imageJ,when you hit the contrast "Auto" button 
once, it sets the percentiles as 1% and 99%.
When you hit it a second time, it increases the stringency by setting the percentiles
to 2% and 98%. As you hit the button more, it brings the bounds closer together
thereby excluding more dark information at the bottom of the histogram and more
bright information at the top of the histogram, emphasizing information in the 
middle, which is usually where image information lies. In cosmx images, reporter
signal usually lies in the middle with background and autofluorescence at the
bottom of the histogram and fiducials and other anomalies (RBCs) at the top.
Contrasts are best set at least by channel, since each channel has different
autofluorescence and fiducial brightness. The settings below are what I am
finding works best for our current samples and channels, but you should adjust
them as you see fit. 

When this parameter is provided, the pptgenerator will
determine a strategy to broadcast it to all of the images within a single display
group. Generally, if you provide a single contrast bound (min%,max%), it will
use that bound to calculate contrast settings for all images, though the actual
contrast bound will be applied to images only within contrast groups. If you provide
6 contrast bounds and there are 48 images in each display group, it will broadcast
those 6 values over the remaining 42 images in a gridlike manner. If you plan to
get the most out of the customization, it's best to match the number of contrast
bounds provided to the fastest iterating display groups, so the last argument
passed to set_groups(), or the last two arguments, or the last three.. etc.
You can also provide a unique contrast bound to every image in the display group,
but that's probably overkill
"""

# contrast_data = [
#     (1,99.5), #BB
#     (1,97), #GG
#     (1,97), #YY
#     (1,99), #RR
#     (0,99), #RN
#     (1,99.5) #BG
#     ]

contrast_data =[(1,99)]

#%% Finally, initialize the ppt generator

#pass the variables and objects defined above
pptg = ppt.pptgenerator(ppt_output_path, cstr, params, contrast_data)

#point the ppt generator to the local template path
pptg.set_template_path(template_path)

#%% Verify the parameters
"""
It can be useful to check some of the automatically generated parameters and
other attributes of the ppt generator to make sure everything looks good. The
most valuable ones are listed here
"""


print('Number of images:', pptg.nims)
print('Grid shape:', pptg.gridshape)

print('Column labels:', pptg.collabels)
print('Row labels:', pptg.rowlabels)

print('Display group:', pptg.cstr.groups)
print('Contrast group:', pptg.contrast_group)

rt = pptg.estimate_run_time()

#%% Verify the parameters part 2
"""
The specifics of how images are arranged into slides and then into grids is
held in the following attributes
"""

print('Show the code layout')
for k,v in pptg.codelist.items():
    print(k)
    for code in v:
        print('\t',str(code))

print('Show the figure layout')
print(pptg.cstr.groups)
for v in pptg.fig_layout:
    print(v)

#%% Generate the slides!
"""
Finally, time to generate the slides. This is the part that takes the longest, 
so be sure that everything set how you want and fire it up. First, the ppt generator
will load images and calculate contrast settings according to the user defined parameters.
After loading the images once, it will make a local copy with the same name as
the ppt file to streamline reloading the image later.
"""

pptg.generate_first() 
pptg.generate()

#%% Cleanup
"""
If you want rerun the ppt generator with significantly different parameters, or
just to save on space, I recommend calling the following method to remove the 
local cache of image data
"""

pptg.remove_temp_dir()

"""
That's it for the tutorial! more info and explanation is available in the
examples. For questions, bug reports, or feature requests, do not hesitate to contact
me at gtremel@nanostring.com and I will help as best I can.
"""


#%% Appendix
"""
Locating run directory
Sooo it's not always guaranteed that the cstringmanager object will correctly
identify the run directory. if you have an issue running the set_dir command,
you can simply assign the necessary parameters directly as shown below.
"""
#parent directory of experiment. i.e. E:\Run1000_my_run
cstr.expdir = exppath
#The datestamp where image data is located
cstr.date = '20240322'
#the timestamp where image data is located
cstr.time = '200553'