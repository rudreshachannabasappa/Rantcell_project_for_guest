import os
import sys
# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.library import *

keys_to_check = ['Map_View', 'Graph_View', 'Exports', 'PDF Data Export and Validation',
                 'Table Summary Export Validation', 'NW Freeze',
                 'Combined Export Data Validation', 'Individual Popup window data validation',
                 'NQC table data validation', 'Default Settings', 'Change Settings',"Group"]
androidtest = side_bar_to_run_for_androidtest(keys_to_check)

print("androidtest",androidtest)