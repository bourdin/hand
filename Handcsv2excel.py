#!/usr/bin/env python
import numpy as np
import hand
import sys

def parse(args=None):
    import argparse
    ### Get options from the command line
    parser = argparse.ArgumentParser(description='Plot energy evolution for VarFracQS.')
    parser.add_argument('inputfile',type=argparse.FileType('r'),help='Input file',default=sys.stdin)
    parser.add_argument('outputfile',help='Output file')

    parser.add_argument('-f','--firstframe',help='first frame number',type=int,default=0)
    parser.add_argument('-l','--lastframe',help='last frame number',type=int,default=None)
    parser.add_argument('-p',"--plot",default=False,action="store_true",help="draw the hand")
    parser.add_argument("--force",action="store_true",default=False,help="Overwrite existing files without prompting")
    return parser.parse_args()

def confirm(prompt=None, resp=False):
    """prompts for yes or no response from the user. Returns True for yes and
    False for no.

    'resp' should be set to the default value assumed by the caller when
    user simply types ENTER.

    >>> confirm(prompt='Create Directory?', resp=True)
    Create Directory? [y]|n: 
    True
    >>> confirm(prompt='Create Directory?', resp=False)
    Create Directory? [n]|y: 
    False
    >>> confirm(prompt='Create Directory?', resp=False)
    Create Directory? [n]|y: y
    True

    """
    
    if prompt is None:
        prompt = 'Confirm'

    if resp:
        prompt = '%s [%s]|%s: ' % (prompt, 'y', 'n')
    else:
        prompt = '%s [%s]|%s: ' % (prompt, 'n', 'y')
        
    while True:
        ans = raw_input(prompt)
        if not ans:
            return resp
        if ans not in ['y', 'Y', 'n', 'N']:
            print 'please enter y or n.'
            continue
        if ans == 'y' or ans == 'Y':
            return True
        if ans == 'n' or ans == 'N':
            return False

def main():
    import hand
    import os
    
    options = parse()
    if options.lastframe == None:
        options.lastframe = 1000000

    if  os.path.exists(options.outputfile):
        if options.force:
            os.remove(options.outputfile)
        else:
            if confirm("Excel file {0} already exists. Overwrite?".format(options.outputfile)):
                os.remove(options.outputfile)
            else:
                print '\n\t{0} was NOT generated.\n'.format(options.outputfile)
                return -1

            
    H = hand.Hand()
    book = H.excelWriteHeaders()

    for f in range(options.firstframe,options.lastframe+1):
        try:
           H.getFrameFromTXT(options.inputfile,frame=f)
           H.excelWriteFrame(book,frame=f)
        except ValueError:
           break
    options.inputfile.close()
    book.save(options.outputfile)
    
if __name__ == "__main__":
    sys.exit(main())
