#------------------------------------------------------------------------------------------#
# Author: Adil Zafar Khan
# Last Edit Date: 12/22/2021
# Description:
"""
    ClassBuilder is a class that contains all the module functions that execute the program.
    It can be viewed as the program manager and every class is finally initialized and
    called from here.
"""
#------------------------------------------------------------------------------------------#

import sys
from Interface import Interface


def main():
    
    frame = Interface()

if __name__ == '__main__':
    
    sys.exit(main())
    
