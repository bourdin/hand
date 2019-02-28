#!/usr/bin/env python3
import numpy as np
import hand
import sys

def parse(args=None):
    import argparse
    ### Get options from the command line
    parser = argparse.ArgumentParser(description='Compute angles of motion in an excel spreadsheet')
    parser.add_argument('inputfile',help='excel spreadsheet file',default=None)
    parser.add_argument('outputfile',help='excel spreadsheet file',default=None)

    parser.add_argument('-s','--sheet',help='Index of the seet to read from',type=int,default=0)
    parser.add_argument('--header',help='Number of header rows',type=int,default=1)
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
        ans = input(prompt)
        if not ans:
            return resp
        if ans not in ['y', 'Y', 'n', 'N']:
            print('please enter y or n.')
            continue
        if ans == 'y' or ans == 'Y':
            return True
        if ans == 'n' or ans == 'N':
            return False

def excelAddRanges(rangeSheet):
    import xlwt
    rangeSheet.write(0,0,'Index')
    rangeSheet.write(1,0,'Abduction Min')
    rangeSheet.write(1,1,xlwt.Formula("MIN('Index'!A1:A65536)"))
    rangeSheet.write(2,0,'Abduction Max')
    rangeSheet.write(2,1,xlwt.Formula("MAX('Index'!A1:A65536)"))
    rangeSheet.write(3,0,'Abduction Range')
    rangeSheet.write(3,1,xlwt.Formula("MAX('Index'!A1:A65536)-MIN('Index'!A1:A65536)"))

    rangeSheet.write(4,0,'MCP Flexion Min')
    rangeSheet.write(4,1,xlwt.Formula("MIN('Index'!B1:B65536)"))
    rangeSheet.write(5,0,'MCP Flexion Max')
    rangeSheet.write(5,1,xlwt.Formula("MAX('Index'!B1:B65536)"))
    rangeSheet.write(6,0,'MCP Flexion Range')
    rangeSheet.write(6,1,xlwt.Formula("MAX('Index'!B1:B65536)-MIN('Index'!B1:B65536)"))

    rangeSheet.write(7,0,'PIP Flexion Min')
    rangeSheet.write(7,1,xlwt.Formula("MIN('Index'!C1:C65536)"))
    rangeSheet.write(8,0,'PIP Flexion Max')
    rangeSheet.write(8,1,xlwt.Formula("MAX('Index'!C1:C65536)"))
    rangeSheet.write(9,0,'PIP Flexion Range')
    rangeSheet.write(9,1,xlwt.Formula("MAX('Index'!C1:C65536)-MIN('Index'!C1:C65536)"))

    rangeSheet.write(10,0,'DIP Flexion Min')
    rangeSheet.write(10,1,xlwt.Formula("MIN('Index'!D1:D65536)"))
    rangeSheet.write(11,0,'DIP Flexion Max')
    rangeSheet.write(11,1,xlwt.Formula("MAX('Index'!D1:D65536)"))
    rangeSheet.write(12,0,'DIP Flexion Range')
    rangeSheet.write(12,1,xlwt.Formula("MAX('Index'!D1:D65536)-MIN('Index'!D1:D65536)"))

    rangeSheet.write(0,2,'Middle')
    rangeSheet.write(1,2,'Abduction Min')
    rangeSheet.write(1,3,xlwt.Formula("MIN('Middle'!A1:A65536)"))
    rangeSheet.write(2,2,'Abduction Max')
    rangeSheet.write(2,3,xlwt.Formula("MAX('Middle'!A1:A65536)"))
    rangeSheet.write(3,2,'Abduction Range')
    rangeSheet.write(3,3,xlwt.Formula("MAX('Middle'!A1:A65536)-MIN('Middle'!A1:A65536)"))

    rangeSheet.write(4,2,'MCP Flexion Min')
    rangeSheet.write(4,3,xlwt.Formula("MIN('Middle'!B1:B65536)"))
    rangeSheet.write(5,2,'MCP Flexion Max')
    rangeSheet.write(5,3,xlwt.Formula("MAX('Middle'!B1:B65536)"))
    rangeSheet.write(6,2,'MCP Flexion Range')
    rangeSheet.write(6,3,xlwt.Formula("MAX('Middle'!B1:B65536)-MIN('Middle'!B1:B65536)"))

    rangeSheet.write(7,2,'PIP Flexion Min')
    rangeSheet.write(7,3,xlwt.Formula("MIN('Middle'!C1:C65536)"))
    rangeSheet.write(8,2,'PIP Flexion Max')
    rangeSheet.write(8,3,xlwt.Formula("MAX('Middle'!C1:C65536)"))
    rangeSheet.write(9,2,'PIP Flexion Range')
    rangeSheet.write(9,3,xlwt.Formula("MAX('Middle'!C1:C65536)-MIN('Middle'!C1:C65536)"))

    rangeSheet.write(10,2,'DIP Flexion Min')
    rangeSheet.write(10,3,xlwt.Formula("MIN('Middle'!D1:D65536)"))
    rangeSheet.write(11,2,'DIP Flexion Max')
    rangeSheet.write(11,3,xlwt.Formula("MAX('Middle'!D1:D65536)"))
    rangeSheet.write(12,2,'DIP Flexion Range')
    rangeSheet.write(12,3,xlwt.Formula("MAX('Middle'!D1:D65536)-MIN('Middle'!D1:D65536)"))

    rangeSheet.write(0,4,'Ring')
    rangeSheet.write(1,4,'Abduction Min')
    rangeSheet.write(1,5,xlwt.Formula("MIN('Ring'!A1:A65536)"))
    rangeSheet.write(2,4,'Abduction Max')
    rangeSheet.write(2,5,xlwt.Formula("MAX('Ring'!A1:A65536)"))
    rangeSheet.write(3,4,'Abduction Range')
    rangeSheet.write(3,5,xlwt.Formula("MAX('Ring'!A1:A65536)-MIN('Ring'!A1:A65536)"))

    rangeSheet.write(4,4,'MCP Flexion Min')
    rangeSheet.write(4,5,xlwt.Formula("MIN('Ring'!B1:B65536)"))
    rangeSheet.write(5,4,'MCP Flexion Max')
    rangeSheet.write(5,5,xlwt.Formula("MAX('Ring'!B1:B65536)"))
    rangeSheet.write(6,4,'MCP Flexion Range')
    rangeSheet.write(6,5,xlwt.Formula("MAX('Ring'!B1:B65536)-MIN('Ring'!B1:B65536)"))

    rangeSheet.write(7,4,'PIP Flexion Min')
    rangeSheet.write(7,5,xlwt.Formula("MIN('Ring'!C1:C65536)"))
    rangeSheet.write(8,4,'PIP Flexion Max')
    rangeSheet.write(8,5,xlwt.Formula("MAX('Ring'!C1:C65536)"))
    rangeSheet.write(9,4,'PIP Flexion Range')
    rangeSheet.write(9,5,xlwt.Formula("MAX('Ring'!C1:C65536)-MIN('Ring'!C1:C65536)"))

    rangeSheet.write(10,4,'DIP Flexion Min')
    rangeSheet.write(10,5,xlwt.Formula("MIN('Ring'!D1:D65536)"))
    rangeSheet.write(11,4,'DIP Flexion Max')
    rangeSheet.write(11,5,xlwt.Formula("MAX('Ring'!D1:D65536)"))
    rangeSheet.write(12,4,'DIP Flexion Range')
    rangeSheet.write(12,5,xlwt.Formula("MAX('Ring'!D1:D65536)-MIN('Ring'!D1:D65536)"))

    rangeSheet.write(0,6,'Little')
    rangeSheet.write(1,6,'Abduction Min')
    rangeSheet.write(1,7,xlwt.Formula("MIN('Little'!A1:A65536)"))
    rangeSheet.write(2,6,'Abduction Max')
    rangeSheet.write(2,7,xlwt.Formula("MAX('Little'!A1:A65536)"))
    rangeSheet.write(3,6,'Abduction Range')
    rangeSheet.write(3,7,xlwt.Formula("MAX('Little'!A1:A65536)-MIN('Little'!A1:A65536)"))

    rangeSheet.write(4,6,'MCP Flexion Min')
    rangeSheet.write(4,7,xlwt.Formula("MIN('Little'!B1:B65536)"))
    rangeSheet.write(5,6,'MCP Flexion Max')
    rangeSheet.write(5,7,xlwt.Formula("MAX('Little'!B1:B65536)"))
    rangeSheet.write(6,6,'MCP Flexion Range')
    rangeSheet.write(6,7,xlwt.Formula("MAX('Little'!B1:B65536)-MIN('Little'!B1:B65536)"))

    rangeSheet.write(7,6,'PIP Flexion Min')
    rangeSheet.write(7,7,xlwt.Formula("MIN('Little'!C1:C65536)"))
    rangeSheet.write(8,6,'PIP Flexion Max')
    rangeSheet.write(8,7,xlwt.Formula("MAX('Little'!C1:C65536)"))
    rangeSheet.write(9,6,'PIP Flexion Range')
    rangeSheet.write(9,7,xlwt.Formula("MAX('Little'!C1:C65536)-MIN('Little'!C1:C65536)"))

    rangeSheet.write(10,6,'DIP Flexion Min')
    rangeSheet.write(10,7,xlwt.Formula("MIN('Little'!D1:D65536)"))
    rangeSheet.write(11,6,'DIP Flexion Max')
    rangeSheet.write(11,7,xlwt.Formula("MAX('Little'!D1:D65536)"))
    rangeSheet.write(12,6,'DIP Flexion Range')
    rangeSheet.write(12,7,xlwt.Formula("MAX('Little'!D1:D65536)-MIN('Little'!D1:D65536)"))

    rangeSheet.write(0,8,'Thumb')
    rangeSheet.write(1,8,'Abduction Min')
    rangeSheet.write(1,9,xlwt.Formula("MIN('Thumb'!A1:A65536)"))
    rangeSheet.write(2,8,'Abduction Max')
    rangeSheet.write(2,9,xlwt.Formula("MAX('Thumb'!A1:A65536)"))
    rangeSheet.write(3,8,'Abduction Range')
    rangeSheet.write(3,9,xlwt.Formula("MAX('Thumb'!A1:A65536)-MIN('Thumb'!A1:A65536)"))

    rangeSheet.write(4,8,'CMC Flexion Min')
    rangeSheet.write(4,9,xlwt.Formula("MIN('Thumb'!B1:B65536)"))
    rangeSheet.write(5,8,'CMC Flexion Max')
    rangeSheet.write(5,9,xlwt.Formula("MAX('Thumb'!B1:B65536)"))
    rangeSheet.write(6,8,'CMC Flexion Range')
    rangeSheet.write(6,9,xlwt.Formula("MAX('Thumb'!B1:B65536)-MIN('Thumb'!B1:B65536)"))

    rangeSheet.write(7,8,'MCP Flexion Min')
    rangeSheet.write(7,9,xlwt.Formula("MIN('Thumb'!C1:C65536)"))
    rangeSheet.write(8,8,'MCP Flexion Max')
    rangeSheet.write(8,9,xlwt.Formula("MAX('Thumb'!C1:C65536)"))
    rangeSheet.write(9,8,'MCP Flexion Range')
    rangeSheet.write(9,9,xlwt.Formula("MAX('Thumb'!C1:C65536)-MIN('Thumb'!C1:C65536)"))

    rangeSheet.write(10,8,'PIP Flexion Min')
    rangeSheet.write(10,9,xlwt.Formula("MIN('Thumb'!D1:D65536)"))
    rangeSheet.write(11,8,'PIP Flexion Max')
    rangeSheet.write(11,9,xlwt.Formula("MAX('Thumb'!D1:D65536)"))
    rangeSheet.write(12,8,'PIP Flexion Range')
    rangeSheet.write(12,9,xlwt.Formula("MAX('Thumb'!D1:D65536)-MIN('Thumb'!D1:D65536)"))

def main():
    import hand
    import os
    import xlrd,xlwt
    
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
                print('\n\t{0} was NOT generated.\n'.format(options.outputfile))
                return -1

    inbook = xlrd.open_workbook(options.inputfile,on_demand=True)
    sheetNum = inbook.sheet_by_index(options.sheet)

    headers = sheetNum.row(options.header-1)
    labels = [s.value.replace('.','').lower().rstrip()[1:] for s in headers]

    outbook = xlwt.Workbook()
    if 'Index' in inbook.sheet_names():
        print('{0} already has a sheet name {1}, delete it and restart'.format(options.inputfile,'Index'))
    else:
        indexSheet  = outbook.add_sheet('Index')
        indexSheet.write(0,0,'Abduction')
        indexSheet.write(0,1,'MCP Flexion')
        indexSheet.write(0,2,'PIP Flexion')
        indexSheet.write(0,3,'DIP Flexion')

    if 'Middle' in inbook.sheet_names():
        print('{0} already has a sheet name {1}, delete it and restart'.format(options.inputfile,'Middle'))
    else:
        middleSheet = outbook.add_sheet('Middle')
        middleSheet.write(0,0,'Abduction')
        middleSheet.write(0,1,'MCP Flexion')
        middleSheet.write(0,2,'PIP Flexion')
        middleSheet.write(0,3,'DIP Flexion')

    if 'Ring' in inbook.sheet_names():
        print('{0} already has a sheet name {1}, delete it and restart'.format(options.inputfile,'Ring'))
    else:
        ringSheet   = outbook.add_sheet('Ring')
        ringSheet.write(0,0,'Abduction')
        ringSheet.write(0,1,'MCP Flexion')
        ringSheet.write(0,2,'PIP Flexion')
        ringSheet.write(0,3,'DIP Flexion')

    if 'Little' in inbook.sheet_names():
        print('{0} already has a sheet name {1}, delete it and restart'.format(options.inputfile,'Little'))
    else:
        littleSheet = outbook.add_sheet('Little')
        littleSheet.write(0,0,'Abduction')
        littleSheet.write(0,1,'MCP Flexion')
        littleSheet.write(0,2,'PIP Flexion')
        littleSheet.write(0,3,'DIP Flexion')

    if 'Thumb' in inbook.sheet_names():
        print('{0} already has a sheet name {1}, delete it and restart'.format(options.inputfile,'Thumb'))
    else:
        thumbSheet  = outbook.add_sheet('Thumb')
        thumbSheet.write(0,0,'Abduction')
        thumbSheet.write(0,1,'CMC Flexion')
        thumbSheet.write(0,2,'MCP Flexion')
        thumbSheet.write(0,3,'PIP Flexion')

    if 'Range' in inbook.sheet_names():
        print('{0} already has a sheet name {1}, delete it and restart'.format(options.inputfile,'Range'))
    else:
        rangeSheet  = outbook.add_sheet('Range')
        excelAddRanges(rangeSheet)

    print (sheetNum.nrows)
    H = hand.Hand()
    for f in range(options.header+options.firstframe,min(options.lastframe+options.header,sheetNum.nrows)):
        # read frame
        frame = sheetNum.row(f)
        H.markers = np.array([float(i.value) for i in frame[-26*3:]]).reshape([26,3])

        # compute hand angles
        H.RadProxArm = H.markers[labels.index('radproxarm_x')//3,:] 
        H.UlnProxArm = H.markers[labels.index('ulnproxarm_x')//3,:] 
        H.RadDisArm  = H.markers[labels.index('raddisarm_x')//3,:] 
        H.UlnDisArm  = H.markers[labels.index('ulndisarm_x')//3,:] 

        H.ThumbCMC   = H.markers[labels.index('thumbcmc_x')//3,:] 
        H.ThumbMCP   = H.markers[labels.index('thumbmcp_x')//3,:] 
        H.ThumbPIP   = H.markers[labels.index('thumbpip_x')//3,:] 
        H.ThumbTIP   = H.markers[labels.index('thumbtip_x')//3,:] 

        H.IndexCMC   = H.markers[labels.index('indexcmc_x')//3,:] 
        H.IndexMCP   = H.markers[labels.index('indexmcp_x')//3,:] 
        H.IndexPIP   = H.markers[labels.index('indexpip_x')//3,:] 
        H.IndexDIP   = H.markers[labels.index('indexdip_x')//3,:] 
        H.IndexTIP   = H.markers[labels.index('indextip_x')//3,:] 

        H.MiddleMCP  = H.markers[labels.index('middlemcp_x')//3,:] 
        H.MiddlePIP  = H.markers[labels.index('middlepip_x')//3,:] 
        H.MiddleDIP  = H.markers[labels.index('middledip_x')//3,:] 
        H.MiddleTIP  = H.markers[labels.index('middletip_x')//3,:] 

        H.RingMCP    = H.markers[labels.index('ringmcp_x')//3,:] 
        H.RingPIP    = H.markers[labels.index('ringpip_x')//3,:] 
        H.RingDIP    = H.markers[labels.index('ringdip_x')//3,:] 
        H.RingTIP    = H.markers[labels.index('ringtip_x')//3,:] 

        H.LittleCMC  = H.markers[labels.index('littlecmc_x')//3,:] 
        H.LittleMCP  = H.markers[labels.index('littlemcp_x')//3,:] 
        H.LittlePIP  = H.markers[labels.index('littlepip_x')//3,:] 
        H.LittleDIP  = H.markers[labels.index('littledip_x')//3,:] 
        H.LittleTIP  = H.markers[labels.index('littletip_x')//3,:] 

        H.CMCVM = (H.IndexCMC + H.LittleCMC)/2.
        H.computeAngles()

        # Write
        indexSheet  = outbook.get_sheet('Index')
        indexSheet.write (f,0,H.IndexAbductionAngle)
        indexSheet.write (f,1,H.IndexMCPFlexionAngle)
        indexSheet.write (f,2,H.IndexPIPFlexionAngle)
        indexSheet.write (f,3,H.IndexDIPFlexionAngle)

        middleSheet  = outbook.get_sheet('Middle')
        middleSheet.write (f,0,H.MiddleAbductionAngle)
        middleSheet.write (f,1,H.MiddleMCPFlexionAngle)
        middleSheet.write (f,2,H.MiddlePIPFlexionAngle)
        middleSheet.write (f,3,H.MiddleDIPFlexionAngle)

        
        ringSheet = outbook.get_sheet('Ring')
        ringSheet.write (f,0,H.RingAbductionAngle)
        ringSheet.write (f,1,H.RingMCPFlexionAngle)
        ringSheet.write (f,2,H.RingPIPFlexionAngle)
        ringSheet.write (f,3,H.RingDIPFlexionAngle)

        littleSheet = outbook.get_sheet('Little')
        littleSheet.write (f,0,H.LittleAbductionAngle)
        littleSheet.write (f,1,H.LittleMCPFlexionAngle)
        littleSheet.write (f,2,H.LittlePIPFlexionAngle)
        littleSheet.write (f,3,H.LittleDIPFlexionAngle)

        thumbSheet = outbook.get_sheet('Thumb')
        thumbSheet.write (f,0,H.ThumbAbductionAngle)
        thumbSheet.write (f,1,H.ThumbCMCFlexionAngle)
        thumbSheet.write (f,2,H.ThumbMCPFlexionAngle)
        thumbSheet.write (f,3,H.ThumbPIPFlexionAngle)

    outbook.save(options.outputfile)
    return 0
    
if __name__ == "__main__":
    sys.exit(main())
