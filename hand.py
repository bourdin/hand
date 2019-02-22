import numpy as np
def angle(v1,v2):
    v1Norm = np.sqrt(np.dot(v1,v1))
    v2Norm = np.sqrt(np.dot(v2,v2))
    return np.arccos(np.dot(v1,v2)/v1Norm/v2Norm) * 180. / np.pi

def vectorProjection(v,d):
    '''Compute the projection of v along the direction d'''
    dNorm = np.sqrt(np.dot(d,d))
    return np.dot(v,d) * v / dNorm
       
def planarProjection(v,n):
    '''Compute the projection of v on the plane orthogonal to n'''
    return v - vectorProjection(v,n)

def outOfPlaneAngle(v1,v2,v):
    '''compute the angle between the vector v and the plane defined by v1 and v2 
       assuming direct orientation.
       if v1 and v2 define the plane of the hand and v a finger, this is the flexion angle.'''
    n = np.cross(v1,v2)
    return 90 - np.arccos(np.dot(n,v)/np.sqrt(np.dot(v,v))/np.sqrt(np.dot(n,n))) * 180 / np.pi 
    
def inPlaneAngle(v1,v2,v):
    '''compute the angle between the vector projection of v on the plane defined by v1 and v2 
       assuming direct orientation and v1.
       If v1 and v2 defines the plane of the hand and v a vector, this is the abduction angle. '''
    n = np.cross(v1,v2)
    vv = planarProjection(v,n)
    theta = np.arccos(np.dot(v2,vv)/np.sqrt(np.dot(v2,v2))/np.sqrt(np.dot(vv,vv))) * 180 / np.pi 
    if np.dot(v2,vv) >= 0:
        return theta
    else:
        return -theta

def getC3dPointsOrdering(file):
    import c3d
    labels = c3d.Reader(file).point_labels
    order = ''
    #print(labels.index('R.Index.TIP'))
    #print(labels.index('R.Index.DIP'))
    #print order

class Hand:
    ''' 3d representation of a hand'''
    
    def __init__(self):
        ''' Initializes a right hand object'''
        markers = np.zeros([26,3])
        self.RadProxArm = 0.
        self.UlnProxArm = 0.
        self.RadDisArm  = 0.
        self.UlnDisArm  = 0.
        
        self.ThumbCMC   = 0.
        self.ThumbMCP   = 0.
        self.ThumbPIP   = 0.
        self.ThumbTIP   = 0.

        self.IndexCMC   = 0.
        self.IndexMCP   = 0.
        self.IndexPIP   = 0.
        self.IndexDIP   = 0.
        self.IndexTIP   = 0.

        self.MiddleMCP  = 0.
        self.MiddlePIP  = 0.
        self.MiddleDIP  = 0.
        self.MiddleTIP  = 0.

        self.RingMCP    = 0.
        self.RingPIP    = 0.
        self.RingDIP    = 0.
        self.RingTIP    = 0.

        self.LittleCMC  = 0.
        self.LittleMCP  = 0.
        self.LittlePIP  = 0.
        self.LittleDIP  = 0.
        self.LittleTIP  = 0.
        
        self.CMCVM      = 0.

        self.IndexAbductionAngle  = 0.
        self.IndexMCPFlexionAngle = 0.
        self.IndexPIPFlexionAngle = 0.
        self.IndexDIPFlexionAngle = 0.
        
        self.MiddleAbductionAngle  = 0.
        self.MiddleMCPFlexionAngle = 0.
        self.MiddlePIPFlexionAngle = 0.
        self.MiddleDIPFlexionAngle = 0.
        
        self.RingAbductionAngle  = 0.
        self.RingMCPFlexionAngle = 0.
        self.RingPIPFlexionAngle = 0.
        self.RingDIPFlexionAngle = 0.
        
        self.LittleAbductionAngle  = 0.
        self.LittleMCPFlexionAngle = 0.
        self.LittlePIPFlexionAngle = 0.
        self.LittleDIPFlexionAngle = 0.
        
        self.ThumbAbductionAngle  = 0.
        self.ThumbCMCFlexionAngle = 0.
        self.ThumbMCPFlexionAngle = 0.
        self.ThumbPIPFlexionAngle = 0.
        
    def computeAngles(self):
        ''' Compute flexion and abduction angle from markers position'''
        # The Motion of the Index finger is in relation to the RHP plane
        v1 = self.IndexMCP - self.MiddleMCP
        v2 = self.IndexMCP - self.CMCVM
        v3 = self.IndexMCP - self.IndexCMC

        self.IndexAbductionAngle  = inPlaneAngle(v1,v3,self.IndexPIP - self.IndexMCP)
        # angle between FI and IN in the plane defined by FI and JI
        self.IndexMCPFlexionAngle = outOfPlaneAngle(v1,v2,self.IndexPIP - self.IndexMCP)
        self.IndexPIPFlexionAngle = outOfPlaneAngle(v1,v2,self.IndexDIP - self.IndexPIP) - self.IndexMCPFlexionAngle
        self.IndexDIPFlexionAngle = outOfPlaneAngle(v1,v2,self.IndexTIP - self.IndexDIP) - self.IndexPIPFlexionAngle - self.IndexMCPFlexionAngle

        # The motion of the Middle and Ring fingers are in relation to the MHP plane
        v1 = self.MiddleMCP - self.RingMCP
        v2 = self.MiddleMCP - self.CMCVM
        v3 = self.RingMCP - self.CMCVM

        self.MiddleAbductionAngle  = inPlaneAngle(v1,v2, self.MiddlePIP - self.MiddleMCP)
        self.MiddleMCPFlexionAngle = outOfPlaneAngle(v1,v2,self.MiddlePIP - self.MiddleMCP)
        self.MiddlePIPFlexionAngle = outOfPlaneAngle(v1,v2,self.MiddleDIP - self.MiddlePIP) - self.MiddleMCPFlexionAngle
        self.MiddleDIPFlexionAngle = outOfPlaneAngle(v1,v2,self.MiddleTIP - self.MiddleDIP) - self.MiddlePIPFlexionAngle - self.MiddleMCPFlexionAngle

        self.RingAbductionAngle  = inPlaneAngle(v1,v3, self.RingPIP - self.RingMCP)
        self.RingMCPFlexionAngle = outOfPlaneAngle(v1,v2,self.RingPIP - self.RingMCP)
        self.RingPIPFlexionAngle = outOfPlaneAngle(v1,v2,self.RingDIP - self.RingPIP) - self.RingMCPFlexionAngle
        self.RingDIPFlexionAngle = outOfPlaneAngle(v1,v2,self.RingTIP - self.RingDIP) - self.RingPIPFlexionAngle - self.RingMCPFlexionAngle

        # The motion of the ring finger is in relation to the UHP plane
        v1 = self.RingMCP - self.LittleMCP
        v2 = self.LittleMCP - self.CMCVM
        v3 = self.LittleMCP - self.LittleCMC

        self.LittleAbductionAngle  = inPlaneAngle (v1,v3,self.LittlePIP - self.LittleMCP)
        self.LittleMCPFlexionAngle = outOfPlaneAngle(v1,v2,self.LittlePIP - self.LittleMCP)
        self.LittlePIPFlexionAngle = outOfPlaneAngle(v1,v2,self.LittleDIP - self.LittlePIP) - self.LittleMCPFlexionAngle
        self.LittleDIPFlexionAngle = outOfPlaneAngle(v1,v2,self.LittleTIP - self.LittleDIP) - self.LittlePIPFlexionAngle - self.LittleMCPFlexionAngle
        
        # The Motion of the Thumb is in relation to the RHP plane
        v1 = self.IndexMCP - self.MiddleMCP
        v2 = self.IndexMCP - self.CMCVM
        v3 = self.IndexMCP - self.IndexCMC

        self.ThumbAbductionAngle = outOfPlaneAngle(v1,v2,self.ThumbMCP - self.ThumbCMC)
        
        self.ThumbCMCFlexionAngle = angle(self.ThumbMCP - self.ThumbCMC,self.IndexMCP - self.IndexCMC)
        # In order to find the sign of the angle, we check if the vector FE x EH and the normal to the MHP plane point in teh same direction
        self.ThumbMCPFlexionAngle = angle(self.ThumbMCP - self.ThumbCMC,self.ThumbPIP - self.ThumbMCP)
        nMHP = np.cross(self.MiddleMCP - self.CMCVM,self.RingMCP - self.CMCVM)
        n1 = np.cross(self.ThumbMCP - self.ThumbCMC, self.ThumbPIP - self.ThumbMCP)
        if np.dot(n1,nMHP) <= 0:
            self.ThumbMCPFlexionAngle = -self.ThumbMCPFlexionAngle 

        self.ThumbPIPFlexionAngle = angle(self.ThumbPIP - self.ThumbMCP,self.ThumbTIP - self.ThumbPIP)
        n1 = np.cross(self.ThumbPIP - self.ThumbMCP,self.ThumbTIP - self.ThumbPIP)
        if np.dot(n1,nMHP) <= 0:
            self.ThumbPIPFlexionAngle = -self.ThumbPIPFlexionAngle 


        
    def getFrameFromTXT(self,fp,frame = 0,header = 3,skip=3,sep = ','):
        ''' Initiate a hand object from a frame in a text file'''

        frameFound = False
        fp.seek(0)
        for i, line in enumerate(fp):
            if i == header-1:
                labels = [s.replace('.','').lower().rstrip()[1:] for s in line.split(sep)]
            if i == skip + frame:
                frameFound = True
                self.markers = np.array([float(i) for i in line.split(sep)]).reshape([26,3])

        if frameFound:        
            self.RadProxArm = self.markers[labels.index('radproxarm_x')/3,:] 
            self.UlnProxArm = self.markers[labels.index('ulnproxarm_x')/3,:] 
            self.RadDisArm  = self.markers[labels.index('raddisarm_x')/3,:] 
            self.UlnDisArm  = self.markers[labels.index('ulndisarm_x')/3,:] 

            self.ThumbCMC   = self.markers[labels.index('thumbcmc_x')/3,:] 
            self.ThumbMCP   = self.markers[labels.index('thumbmcp_x')/3,:] 
            self.ThumbPIP   = self.markers[labels.index('thumbpip_x')/3,:] 
            self.ThumbTIP   = self.markers[labels.index('thumbtip_x')/3,:] 

            self.IndexCMC   = self.markers[labels.index('indexcmc_x')/3,:] 
            self.IndexMCP   = self.markers[labels.index('indexmcp_x')/3,:] 
            self.IndexPIP   = self.markers[labels.index('indexpip_x')/3,:] 
            self.IndexDIP   = self.markers[labels.index('indexdip_x')/3,:] 
            self.IndexTIP   = self.markers[labels.index('indextip_x')/3,:] 

            self.MiddleMCP  = self.markers[labels.index('middlemcp_x')/3,:] 
            self.MiddlePIP  = self.markers[labels.index('middlepip_x')/3,:] 
            self.MiddleDIP  = self.markers[labels.index('middledip_x')/3,:] 
            self.MiddleTIP  = self.markers[labels.index('middletip_x')/3,:] 

            self.RingMCP    = self.markers[labels.index('ringmcp_x')/3,:] 
            self.RingPIP    = self.markers[labels.index('ringpip_x')/3,:] 
            self.RingDIP    = self.markers[labels.index('ringdip_x')/3,:] 
            self.RingTIP    = self.markers[labels.index('ringtip_x')/3,:] 

            self.LittleCMC  = self.markers[labels.index('littlecmc_x')/3,:] 
            self.LittleMCP  = self.markers[labels.index('littlemcp_x')/3,:] 
            self.LittlePIP  = self.markers[labels.index('littlepip_x')/3,:] 
            self.LittleDIP  = self.markers[labels.index('littledip_x')/3,:] 
            self.LittleTIP  = self.markers[labels.index('littletip_x')/3,:] 

            self.CMCVM = (self.IndexCMC + self.LittleCMC)/2.
            self.computeAngles()
        else:
            raise ValueError('Cannot find frame {0} in file'.format(frame))
            
    def printAngles(self):
        print('Index:')
        print('\tAbduction:   {0:.2f} \t MCP Flexion:   {1:.2f}'.format(self.IndexAbductionAngle,self.IndexMCPFlexionAngle))
        print('\tPIP Flexion: {0:.2f} \t DIP Flexion:   {1:.2f}'.format(self.IndexPIPFlexionAngle,self.IndexDIPFlexionAngle))
        print('Middle:')
        print('\tAbduction:   {0:.2f} \t MCP Flexion:   {1:.2f}'.format(self.MiddleAbductionAngle,self.MiddleMCPFlexionAngle))
        print('\tPIP Flexion: {0:.2f} \t DIP Flexion:   {1:.2f}'.format(self.MiddlePIPFlexionAngle,self.MiddleDIPFlexionAngle))
        print('Ring:')
        print('\tAbduction:   {0:.2f} \t MCP Flexion:   {1:.2f}'.format(self.RingAbductionAngle,self.RingMCPFlexionAngle))
        print('\tPIP Flexion: {0:.2f} \t DIP Flexion:   {1:.2f}'.format(self.RingPIPFlexionAngle,self.RingDIPFlexionAngle))
        print('Little:')
        print('\tAbduction:   {0:.2f} \t MCP Flexion:   {1:.2f}'.format(self.LittleAbductionAngle,self.LittleMCPFlexionAngle))
        print('\tPIP Flexion: {0:.2f} \t DIP Flexion:   {1:.2f}'.format(self.LittlePIPFlexionAngle,self.LittleDIPFlexionAngle))
        print('Thumb:')
        print('\tAbduction:   {0:.2f} \t CMC Flexion:   {1:.2f}'.format(self.ThumbAbductionAngle,self.ThumbCMCFlexionAngle))
        print('\tMCP Flexion: {0:.2f} \t PIP Flexion:   {1:.2f}'.format(self.ThumbMCPFlexionAngle,self.ThumbPIPFlexionAngle))

    def printThumbAnglesHeader(self):
        print("Thumb\nAbduction\tCMC Flexion\tMCP Flexion\tPIP Flexion")
        
    def printThumbAnglesShort(self):
        print("{0:.2f}\t\t{1:.2f}\t\t{2:.2f}\t\t{3:.2f}".format(self.ThumbAbductionAngle,self.ThumbCMCFlexionAngle ,self.ThumbMCPFlexionAngle,self.ThumbPIPFlexionAngle ))
        
    def excelWriteHeaders(self):
        import xlwt
        book = xlwt.Workbook()
        indexSheet  = book.add_sheet('Index')
        middleSheet = book.add_sheet('Middle')
        ringSheet   = book.add_sheet('Ring')
        littleSheet = book.add_sheet('Little')
        thumbSheet  = book.add_sheet('Thumb')
        indexSheet.write(0,0,'Abduction')
        indexSheet.write(0,1,'MCP Flexion')
        indexSheet.write(0,2,'PIP Flexion')
        indexSheet.write(0,3,'DIP Flexion')
        middleSheet.write(0,0,'Abduction')
        middleSheet.write(0,1,'MCP Flexion')
        middleSheet.write(0,2,'PIP Flexion')
        middleSheet.write(0,3,'DIP Flexion')
        ringSheet.write(0,0,'Abduction')
        ringSheet.write(0,1,'MCP Flexion')
        ringSheet.write(0,2,'PIP Flexion')
        ringSheet.write(0,3,'DIP Flexion')
        littleSheet.write(0,0,'Abduction')
        littleSheet.write(0,1,'MCP Flexion')
        littleSheet.write(0,2,'PIP Flexion')
        littleSheet.write(0,3,'DIP Flexion')
        thumbSheet.write(0,0,'Abduction')
        thumbSheet.write(0,1,'CMC Flexion')
        thumbSheet.write(0,2,'MCP Flexion')
        thumbSheet.write(0,3,'PIP Flexion')
        return book

    def excelWriteFrame(self,book,frame):
        indexSheet  = book.get_sheet('Index')
        indexSheet.write (frame+1,0,self.IndexAbductionAngle)
        indexSheet.write (frame+1,1,self.IndexMCPFlexionAngle)
        indexSheet.write (frame+1,2,self.IndexPIPFlexionAngle)
        indexSheet.write (frame+1,3,self.IndexDIPFlexionAngle)

        middleSheet  = book.get_sheet('Middle')
        middleSheet.write (frame+1,0,self.MiddleAbductionAngle)
        middleSheet.write (frame+1,1,self.MiddleMCPFlexionAngle)
        middleSheet.write (frame+1,2,self.MiddlePIPFlexionAngle)
        middleSheet.write (frame+1,3,self.MiddleDIPFlexionAngle)

        
        ringSheet = book.get_sheet('Ring')
        ringSheet.write (frame+1,0,self.RingAbductionAngle)
        ringSheet.write (frame+1,1,self.RingMCPFlexionAngle)
        ringSheet.write (frame+1,2,self.RingPIPFlexionAngle)
        ringSheet.write (frame+1,3,self.RingDIPFlexionAngle)

        littleSheet = book.get_sheet('Little')
        littleSheet.write (frame+1,0,self.LittleAbductionAngle)
        littleSheet.write (frame+1,1,self.LittleMCPFlexionAngle)
        littleSheet.write (frame+1,2,self.LittlePIPFlexionAngle)
        littleSheet.write (frame+1,3,self.LittleDIPFlexionAngle)

        thumbSheet = book.get_sheet('Thumb')
        thumbSheet.write (frame+1,0,self.ThumbAbductionAngle)
        thumbSheet.write (frame+1,1,self.ThumbCMCFlexionAngle)
        thumbSheet.write (frame+1,2,self.ThumbMCPFlexionAngle)
        thumbSheet.write (frame+1,3,self.ThumbPIPFlexionAngle)

        

    def view(self,figsize=(5,5),markersize=5,linewidth=5):
        import matplotlib.pyplot as plt
        from mpl_toolkits.mplot3d import Axes3D
        from mpl_toolkits.mplot3d.art3d import Poly3DCollection
        
        fig = plt.figure()
        axs = Axes3D(fig)
        cs = ['#5da5da', '#faa43a' , '#60bd68' , '#f17cb0' , '#b2912f' , '#b276b2' , '#decf3f' , '#f15854']

        # Plot all markers
        axs.scatter([self.CMCVM[0],],[self.CMCVM[1],],[self.CMCVM[2],],s=markersize**2,alpha=.75)
        axs.scatter(self.markers[:,0],self.markers[:,1],self.markers[:,2],s=markersize**2,alpha=.75)

        
        # Plot fingers
        Index = np.array([self.IndexMCP,self.IndexPIP,self.IndexDIP,self.IndexTIP])
        axs.plot(Index[:,0],Index[:,1],Index[:,2],linewidth=linewidth,alpha=.75)
        
        Middle = np.array([self.MiddleMCP,self.MiddlePIP,self.MiddleDIP,self.MiddleTIP])
        axs.plot(Middle[:,0],Middle[:,1],Middle[:,2],linewidth=linewidth,alpha=.75)
        
        Ring = np.array([self.RingMCP,self.RingPIP,self.RingDIP,self.RingTIP])
        axs.plot(Ring[:,0],Ring[:,1],Ring[:,2],linewidth=linewidth,alpha=.75)
        
        Little = np.array([self.LittleMCP,self.LittlePIP,self.LittleDIP,self.LittleTIP])
        axs.plot(Little[:,0],Little[:,1],Little[:,2],linewidth=linewidth,alpha=.75)
        
        Thumb = np.array([self.ThumbCMC,self.ThumbMCP,self.ThumbPIP,self.ThumbTIP])
        axs.plot(Thumb[:,0],Thumb[:,1],Thumb[:,2],linewidth=linewidth,alpha=.75)
        
        # plot arm and hand 
        c = 0
        arm = [self.RadProxArm,self.UlnProxArm,self.UlnDisArm,self.RadDisArm,]
        patch = Poly3DCollection([[p for p in arm],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)
        
        c+=1
        c = c%len(cs)
        RHP = [self.MiddleMCP,self.IndexMCP,self.CMCVM]
        patch = Poly3DCollection([[p for p in RHP],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)

        c+=1
        c = c%len(cs)
        MHP = [self.RingMCP,self.MiddleMCP,self.CMCVM]
        patch = Poly3DCollection([[p for p in MHP],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)

        c+=1
        c = c%len(cs)
        UHP = [self.LittleMCP,self.RingMCP,self.CMCVM]
        patch = Poly3DCollection([[p for p in UHP],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)
        
        c+=1
        c = c%len(cs)
        P = [self.IndexMCP,self.IndexCMC,self.CMCVM]
        patch = Poly3DCollection([[p for p in P],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)

        c+=1
        c = c%len(cs)
        P = [self.CMCVM,self.LittleCMC,self.LittleMCP]
        patch = Poly3DCollection([[p for p in P],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)

        c+=1
        c = c%len(cs)
        P = [self.IndexMCP,self.ThumbMCP,self.ThumbCMC,self.IndexCMC,self.IndexMCP]
        patch = Poly3DCollection([[p for p in P],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)

        c+=1
        c = c%len(cs)
        P = [self.RadDisArm,self.UlnDisArm,self.LittleCMC,self.IndexCMC,self.ThumbCMC]
        patch = Poly3DCollection([[p for p in P],])
        patch.set_alpha(.75)
        patch.set_color(cs[c])
        patch.set_edgecolor('k')
        axs.add_collection3d(patch)
        plt.show()
        return fig,axs
    
    def view2D(self,markersize=5):
        import matplotlib.pyplot as plt
        fig = plt.figure()

        # Plot all markers
        plt.scatter([self.CMCVM[0],],[self.CMCVM[1],],s=markersize**2,alpha=.75)
        plt.scatter(self.markers[:,0],self.markers[:,1],s=markersize**2,alpha=.75)

        plt.show()
      