#!/usr/bin/env python
import numpy as np
import hand
import sys

def parse(args=None):
    import argparse
    ### Get options from the command line
    parser = argparse.ArgumentParser(description='Plot energy evolution for VarFracQS.')
    parser.add_argument('inputfile',type=argparse.FileType('r'),nargs='?',help='Input file',default=sys.stdin)
    parser.add_argument('-f','--firstframe',help='first frame number',type=int,default=0)
    parser.add_argument('-l','--lastframe',help='last frame number',type=int,default=0)
    parser.add_argument('-p',"--plot",default=False,action="store_true",help="draw the hand")
    return parser.parse_args()


def main():
    import hand
    options = parse()
    print options
    H = hand.Hand()
    H.printThumbAnglesHeader()
    for f in range(options.firstframe,options.lastframe+1):
        H.getFrameFromTXT(options.inputfile,frame=f)
        H.printThumbAnglesShort()
        if options.plot:
            H.view()
    options.inputfile.close()
    
if __name__ == "__main__":
    sys.exit(main())
