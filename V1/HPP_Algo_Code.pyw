import os
import sys

# Ensure the script can import the main module from the same folder.
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import HPP_Algo_Code
