# -*- coding: utf-8 -*-
"""
Created on Wed Mar 13 17:21:02 2019

@author: N.Srisutthiyakorn
"""

from pptx import Presentation
from pptx.util import Inches
import os
import numpy as np


pp_dir          = r'C:\Users\N.Srisutthiyakorn\OneDrive - Shell\Documents\GitHub\Python_Toolbox_NS\PPToolboxExample'
pp_fnTemplate   = os.path.join(pp_dir, 'ShellTemplate.pptx') 
pp_fnOutput     = os.path.join(pp_dir, 'PPToolboxOutput.pptx')
fig_dir         = r'C:\Users\N.Srisutthiyakorn\OneDrive - Shell\Documents\GitHub\Python_Toolbox_NS\PPToolboxExample\MainFigureFolder'

createPPT(pp_fnTemplate, pp_fnOutput, fig_dir)   