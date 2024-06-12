import os
import sys
# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.library import *

ListofCampaigns_runvalue = Testrun_mode(value= "List of Campaigns")
print("ListofCampaigns_runvalue",ListofCampaigns_runvalue)
