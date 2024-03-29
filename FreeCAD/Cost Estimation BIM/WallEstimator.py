import FreeCAD, math
import numpy as np

class WallEstimator:
    '''
    All units are in mm
    '''
    def __init__(self):
        self.Length = 0
        self.Width = 0
        self.Height = 0
        self.Volume = 0
        self.num_bricks = 0
        self.brick_cost = 0
        self.dry_vol_sand = 0
        self.sand_small_trolley = 0
        self.sand_large_trolley = 0
        self.sand_truck = 0
        self.sand_big_truck = 0
        self.sand_seenkara = 0
        self.num_cement = 0
    
    def wall_cost_estimate(self, obj, brick_unit_cost=12 ,debug=False):
        '''
        Function to estimate cost of a wall
        '''
        # Make it static
        volume_of_wall = obj.Shape.Volume
        volume_of_brick = 9.375*4.875*3.375*25.4**3 # in mm
        
        num_bricks = math.ceil(volume_of_wall/volume_of_brick)
        losses = 5/100
        brick_cost = (num_bricks + num_bricks*losses)* brick_unit_cost
        
        actual_vol_brick = 9*4.5*3*25.4**3
        
        vol_mortar = volume_of_wall - num_bricks*actual_vol_brick
        #print(volume_of_wall/10**6, volume_of_brick*num_bricks/10**6,vol_mortar/10**6)
        dry_vol_mortar = 1.54*vol_mortar
        vol_cement = vol_mortar * 1/5
        vol_sand = vol_mortar * 4/5
        
        dry_vol_sand = vol_sand*1.333
        
        cement_bag = 1.25*(12*25.4)**3
        num_cement = vol_cement/cement_bag
        
        small_trolley = 75*(12*25.4)**3
        large_trolley = 325*(12*25.4)**3
        seenkara = 100*(12*25.4)**3
        truck = 5.5*seenkara
        big_truck = 12*seenkara 
        
        sand_truck = dry_vol_sand/truck
        sand_big_truck = dry_vol_sand/truck
        sand_small_trolley = dry_vol_sand/small_trolley
        sand_large_trolley =dry_vol_sand/large_trolley
        sand_seenkara =dry_vol_sand/seenkara
        
        self.num_bricks += num_bricks
        self.brick_cost += brick_cost
        self.dry_vol_sand += dry_vol_sand
        self.sand_small_trolley += sand_small_trolley
        self.sand_large_trolley += sand_large_trolley
        self.sand_truck += sand_truck
        self.sand_big_truck += sand_big_truck
        self.sand_seenkara += sand_seenkara
        self.num_cement += num_cement
        self.Length +=  obj.Shape.Length
        self.Volume += obj.Shape.Volume
    
    def all_wall_estimate(self):
        '''
        Function to estimate cost of all walls
        '''
        for obj in FreeCAD.ActiveDocument.Objects:
            try:
                if obj.IfcRole == 'Wall':
                    self.wall_cost_estimate(obj)
                elif obj.IfcType == 'Wall':
                    self.wall_cost_estimate(obj)
            except AttributeError:
                pass
        
        return vars(self)
    
    def selected_wall_estimate(self):
        '''
        Function to estimate cost of selected walls
        Don't use the same object twice with this method
        '''
        for obj in FreeCADGui.Selection.getSelection():
            try:
                if obj.IfcRole == 'Wall':
                    self.wall_cost_estimate(obj)
                elif obj.IfcType == 'Wall':
                    self.wall_cost_estimate(obj)
            except AttributeError:
                pass
        
        return vars(self)
    
    def pretty_all_wall_estimate(self):
        '''
        Function to estimate cost of all walls
        '''
        for obj in FreeCAD.ActiveDocument.Objects:
            try:
                if obj.IfcRole == 'Wall':
                    self.wall_cost_estimate(obj)
                elif obj.IfcType == 'Wall':
                    self.wall_cost_estimate(obj)
            except AttributeError:
                pass
            
        print('Number of Bricks :', self.num_bricks)
        print('Cement Bags: ', self.num_cement)
        print('Volume of Sand (cu ft):', self.dry_vol_sand*(10**-3*3.28)**3, 'cubic feet')

WallEstimator().pretty_all_wall_estimate()
WallEstimator().selected_wall_estimate()

