import os
import sys
# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.library import *

datetime_runvalue = Testrun_mode(value="Date and Time")
print("datetime_runvalue",datetime_runvalue)