import os
import sys
# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.library import *

Remotetest_runvalue = Testrun_mode(value="Remote Test")
scheduletest_runvalue = Testrun_mode(value="Schedule test")
continuoustest_runvalue = Testrun_mode(value="Continuous Test")
print("Remotetest_runvalue",Remotetest_runvalue)
print("scheduletest_runvalue",scheduletest_runvalue)
print("continuoustest_runvalue",continuoustest_runvalue)