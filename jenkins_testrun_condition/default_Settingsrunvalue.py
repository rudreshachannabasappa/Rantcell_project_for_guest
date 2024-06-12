import os
import sys
# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.library import *

default_Settingsrunvalue = Testrun_mode(value="Default Settings")
print("Default_Settingsrunvalue",default_Settingsrunvalue)
