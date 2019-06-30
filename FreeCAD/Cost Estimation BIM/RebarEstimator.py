import FreeCAD, math
import numpy as np

class RebarEstimator:
    '''
    All units are in mm
    '''
    def __init__(self):
        self.Length = 0
        self.Width = 0
        self.Height = 0
        self.Volume = 0
        self.mass = 0 # in kg
    
    def cost_estimate(self, obj, brick_unit_cost=12 ,debug=False):
        '''
        Function to estimate cost of a wall
        '''
        # Make it static
        volume_of_wall = obj.Shape.Volume
        print(volume_of_wall)
    
    def all_estimate(self):
        '''
        Function to estimate cost of all walls
        '''
        for obj in FreeCAD.ActiveDocument.Objects:
            try:
                if obj.IfcRole == 'Reinforcing Bar':
                    self.wall_cost_estimate(obj)
                elif obj.IfcType == 'Reinforcing Bar':
                    self.wall_cost_estimate(obj)
            except AttributeError:
                pass
        
        return vars(self)
    
    def selected_estimate(self):
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
    
    def pretty_all_estimate(self):
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

RebarEstimator().all_estimate()


