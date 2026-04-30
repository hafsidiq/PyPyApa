# -*- coding: utf-8 -*-
"""Generates Auto Formwork (15mm panels) separated strictly per element with automatic hole cutouts for intersections."""

import clr
import os
import tempfile

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')

import System.Windows
from System.Windows.Controls import CheckBox
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.Generic import List
from pyrevit import revit, UI, forms

doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# API VERSION COMPATIBILITY
# ==========================================
def get_id_value(id_obj):
    try: return id_obj.Value
    except AttributeError: return id_obj.IntegerValue

# ==========================================
# 1. XAML UI DESIGN (Embedded)
# ==========================================
XAML_CONTENT = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Auto Formwork Pro - Precision Sliced" Width="440" Height="720" 
        WindowStartupLocation="CenterScreen" Background="#F8F9FA" FontFamily="Segoe UI">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="1. Selection Scope" FontWeight="Bold" FontSize="14" Margin="0,0,0,8"/>
        <ComboBox Grid.Row="1" x:Name="cmbScope" Margin="0,0,0,10" SelectedIndex="0" Padding="5">
            <ComboBoxItem Content="Active View Workspace"/>
            <ComboBoxItem Content="Entire Project"/>
            <ComboBoxItem Content="Select By Level"/>
            <ComboBoxItem Content="Manual Selection (Cursor)"/>
        </ComboBox>
        
        <ScrollViewer Grid.Row="2" x:Name="scrollLevels" Visibility="Collapsed" MaxHeight="90" Margin="0,0,0,15" Background="White" BorderBrush="#CED4DA" BorderThickness="1">
            <StackPanel x:Name="pnlLevels" Margin="5"/>
        </ScrollViewer>

        <TextBlock Grid.Row="3" Text="2. Structural Categories to Process" FontWeight="Bold" FontSize="14" Margin="0,0,0,8"/>
        <StackPanel Grid.Row="4" Margin="0,0,0,15">
            <CheckBox x:Name="chkColumns" Content="Structural Columns" IsChecked="True" Margin="0,3"/>
            <CheckBox x:Name="chkBeams" Content="Structural Framing (Beams)" IsChecked="True" Margin="0,3"/>
            <CheckBox x:Name="chkFoundations" Content="Structural Foundations" IsChecked="True" Margin="0,3"/>
            <CheckBox x:Name="chkFloors" Content="Floors (Slabs)" IsChecked="True" Margin="0,3"/>
            <CheckBox x:Name="chkWalls" Content="Walls" IsChecked="True" Margin="0,3"/>
            <CheckBox x:Name="chkStairs" Content="Stairs (Concrete)" IsChecked="True" Margin="0,3"/>
            <CheckBox x:Name="chkGeneric" Content="Generic Models (Concrete)" IsChecked="True" Margin="0,3"/>
        </StackPanel>
        
        <StackPanel Grid.Row="6" Margin="0,0,0,20" Background="#E9ECEF">
            <CheckBox x:Name="chkMTO" Content="Auto-Generate Material Takeoff Schedule" IsChecked="True" Margin="10"/>
        </StackPanel>

        <Button Grid.Row="8" x:Name="btnRun" Content="GENERATE SLICED FORMWORK (15mm)" Height="45" 
                Background="#005A9E" Foreground="White" FontWeight="Bold" FontSize="14" Cursor="Hand"/>
    </Grid>
</Window>
"""

class FormworkUI(forms.WPFWindow):
    def __init__(self, xaml_file_path, doc):
        forms.WPFWindow.__init__(self, xaml_file_path)
        self.ExecuteCode = False
        self.doc = doc
        
        self.levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        self.level_checkboxes = []
        for lvl in self.levels:
            cb = CheckBox()
            cb.Content = lvl.Name
            cb.Tag = lvl.Id
            cb.Margin = System.Windows.Thickness(0, 2, 0, 2)
            self.pnlLevels.Children.Add(cb)
            self.level_checkboxes.append(cb)
            
        self.cmbScope.SelectionChanged += self.scope_changed
        self.btnRun.Click += self.btnRun_Click

    def scope_changed(self, sender, e):
        if self.cmbScope.SelectedIndex == 2: 
            self.scrollLevels.Visibility = System.Windows.Visibility.Visible
        else: 
            self.scrollLevels.Visibility = System.Windows.Visibility.Collapsed

    def btnRun_Click(self, sender, e):
        self.ExecuteCode = True
        self.scope = self.cmbScope.Text
        self.do_columns = self.chkColumns.IsChecked
        self.do_beams = self.chkBeams.IsChecked
        self.do_foundations = self.chkFoundations.IsChecked
        self.do_floors = self.chkFloors.IsChecked
        self.do_walls = self.chkWalls.IsChecked
        self.do_stairs = self.chkStairs.IsChecked
        self.do_generic = self.chkGeneric.IsChecked
        self.generate_mto = self.chkMTO.IsChecked
        self.selected_level_ids = [cb.Tag for cb in self.level_checkboxes if cb.IsChecked]
        self.Close()

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================
class StructuralSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if not elem.Category: return False
        cat_val = int(get_id_value(elem.Category.Id))
        allowed = [
            int(BuiltInCategory.OST_StructuralColumns), int(BuiltInCategory.OST_StructuralFraming), 
            int(BuiltInCategory.OST_StructuralFoundation), int(BuiltInCategory.OST_Floors), 
            int(BuiltInCategory.OST_Walls), int(BuiltInCategory.OST_Stairs), int(BuiltInCategory.OST_GenericModel)
        ]
        return cat_val in allowed
    def AllowReference(self, ref, xyz): 
        return False

def is_element_on_levels(elem, level_ids):
    for p in [BuiltInParameter.FAMILY_BASE_LEVEL_PARAM, BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM, BuiltInParameter.LEVEL_PARAM]:
        param = elem.get_Parameter(p)
        if param and param.AsElementId() in level_ids: 
            return True
    return False

def get_or_create_wood_material(doc):
    for mat in FilteredElementCollector(doc).OfClass(Material).ToElements():
        if "Plywood" in mat.Name or "Kayu" in mat.Name or "Cherry" in mat.Name: 
            return mat.Id
    mat_id = Material.Create(doc, "Formwork - Plywood")
    mat = doc.GetElement(mat_id)
    mat.Color = Color(210, 180, 140) 
    return mat_id

def get_element_solid(element):
    geom_options = Options()
    geom_options.DetailLevel = ViewDetailLevel.Fine
    geom_element = element.get_Geometry(geom_options)
    solids = []
    
    if geom_element:
        for geom_obj in geom_element:
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0: 
                solids.append(geom_obj)
            elif isinstance(geom_obj, GeometryInstance):
                for inst_obj in geom_obj.GetInstanceGeometry():
                    if isinstance(inst_obj, Solid) and inst_obj.Volume > 0: 
                        solids.append(inst_obj)
    
    if not solids: return None
    main_solid = solids[0]
    for s in solids[1:]:
        try: main_solid = BooleanOperationsUtils.ExecuteBooleanOperation(main_solid, s, BooleanOperationsType.Union)
        except: pass
    return main_solid

def get_intersecting_solids(elem, doc, valid_cats):
    """Finds all surrounding structural elements and extracts their solids to be used as 'Cutters'."""
    bb = elem.get_BoundingBox(None)
    if not bb: return []
    
    # Expand bounding box slightly (1 foot) to catch all flush/touching elements
    min_pt = XYZ(bb.Min.X - 1.0, bb.Min.Y - 1.0, bb.Min.Z - 1.0)
    max_pt = XYZ(bb.Max.X + 1.0, bb.Max.Y + 1.0, bb.Max.Z + 1.0)
    outline = Outline(min_pt, max_pt)
    
    bb_filter = BoundingBoxIntersectsFilter(outline)
    nearby_elems = FilteredElementCollector(doc).WherePasses(bb_filter).WhereElementIsNotElementType().ToElements()
    
    invaders = []
    valid_cat_ints = [int(get_id_value(c)) for c in valid_cats]
    
    for n in nearby_elems:
        if n.Id == elem.Id: continue # Skip itself
        if not n.Category: continue
        if int(get_id_value(n.Category.Id)) not in valid_cat_ints: continue
        
        s = get_element_solid(n)
        if s: invaders.append(s)
        
    return invaders

def get_category_name(cat_val):
    cat_names = {
        int(BuiltInCategory.OST_StructuralColumns): "Column",
        int(BuiltInCategory.OST_StructuralFraming): "Beam",
        int(BuiltInCategory.OST_StructuralFoundation): "Foundation",
        int(BuiltInCategory.OST_Floors): "Slab",
        int(BuiltInCategory.OST_Walls): "Wall",
        int(BuiltInCategory.OST_Stairs): "Stairs",
        int(BuiltInCategory.OST_GenericModel): "Generic Model"
    }
    return cat_names.get(cat_val, "Other")

def create_formwork_mto(doc):
    mto_name = "Auto Formwork MTO (Precision Cut)"
    for sch in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements():
        if sch.Name == mto_name: return sch
        
    cat_id = ElementId(BuiltInCategory.OST_GenericModel)
    mto = ViewSchedule.CreateMaterialTakeoff(doc, cat_id)
    mto.Name = mto_name
    fields = mto.Definition.GetSchedulableFields()
    
    f_comments, f_mark, f_mat_name, f_mat_area, f_mat_vol = None, None, None, None, None
    for f in fields:
        name = f.GetName(doc)
        if name == "Comments": f_comments = f
        elif name == "Mark": f_mark = f
        elif name == "Material: Name": f_mat_name = f
        elif name == "Material: Area": f_mat_area = f
        elif name == "Material: Volume": f_mat_vol = f
        
    if f_comments: mto.Definition.AddField(f_comments)
    if f_mark: mto.Definition.AddField(f_mark)  
    if f_mat_name: mto.Definition.AddField(f_mat_name)
    if f_mat_area: 
        area_f = mto.Definition.AddField(f_mat_area)
        area_f.DisplayType = ScheduleFieldDisplayType.Totals
    if f_mat_vol: 
        vol_f = mto.Definition.AddField(f_mat_vol)
        vol_f.DisplayType = ScheduleFieldDisplayType.Totals
    
    try:
        sort_field = ScheduleSortGroupField(f_comments.FieldId)
        mto.Definition.AddSortGroupField(sort_field)
    except: pass
    
    mto.Definition.IsItemized = True 
    return mto

# ==========================================
# 3. MAIN EXECUTION
# ==========================================
def main():
    temp_xaml = tempfile.NamedTemporaryFile(delete=False, suffix=".xaml")
    temp_xaml.write(XAML_CONTENT.encode('utf-8'))
    temp_xaml.close()
    
    ui = FormworkUI(temp_xaml.name, doc)
    ui.ShowDialog()
    os.remove(temp_xaml.name)
    
    if not ui.ExecuteCode: return

    target_cats = []
    if ui.do_columns: target_cats.append(ElementId(BuiltInCategory.OST_StructuralColumns))
    if ui.do_beams: target_cats.append(ElementId(BuiltInCategory.OST_StructuralFraming))
    if ui.do_foundations: target_cats.append(ElementId(BuiltInCategory.OST_StructuralFoundation))
    if ui.do_floors: target_cats.append(ElementId(BuiltInCategory.OST_Floors))
    if ui.do_walls: target_cats.append(ElementId(BuiltInCategory.OST_Walls))
    if ui.do_stairs: target_cats.append(ElementId(BuiltInCategory.OST_Stairs))
    if ui.do_generic: target_cats.append(ElementId(BuiltInCategory.OST_GenericModel))
    
    if not target_cats:
        return forms.alert("Please check at least one structural category.")

    elements = []
    if ui.scope == "Manual Selection (Cursor)":
        try:
            sel_filter = StructuralSelectionFilter()
            references = uidoc.Selection.PickObjects(ObjectType.Element, sel_filter, "Select Elements")
            valid_cats_ints = [int(get_id_value(c)) for c in target_cats]
            elements = [doc.GetElement(ref) for ref in references if int(get_id_value(doc.GetElement(ref).Category.Id)) in valid_cats_ints]
        except: return 
    else:
        for cat_id in target_cats:
            cat_enum = System.Enum.ToObject(BuiltInCategory, int(get_id_value(cat_id)))
            if ui.scope == "Active View Workspace": 
                collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(cat_enum).WhereElementIsNotElementType()
            else: 
                collector = FilteredElementCollector(doc).OfCategory(cat_enum).WhereElementIsNotElementType()
            
            if ui.scope == "Select By Level": 
                elements.extend([e for e in collector.ToElements() if is_element_on_levels(e, ui.selected_level_ids)])
            else: 
                elements.extend(collector.ToElements())

    if not elements: return forms.alert("No elements found.")

    # Execute Geometry Creation
    total_generated = 0
    formwork_thickness = 15.0 / 304.8 
    generic_model_cat = ElementId(BuiltInCategory.OST_GenericModel)

    with revit.Transaction("Generate Precision Cut Formwork"):
        wood_mat_id = get_or_create_wood_material(doc)
        try: solid_opts = SolidOptions(wood_mat_id, ElementId.InvalidElementId)
        except: solid_opts = None

        with forms.ProgressBar(title='Cutting Holes & Generating Formwork...', step=100) as pb:
            for i, elem in enumerate(elements):
                cat_val = int(get_id_value(elem.Category.Id))
                category_string = get_category_name(cat_val)
                
                solid = get_element_solid(elem)
                if not solid: continue
                
                # GET ALL SURROUNDING INTERSECTING ELEMENTS
                invading_solids = get_intersecting_solids(elem, doc, target_cats)
                
                element_formwork_panels = List[GeometryObject]()
                
                for face in solid.Faces:
                    if not isinstance(face, PlanarFace): continue
                    normal = face.FaceNormal
                    
                    is_foundation = (cat_val == int(BuiltInCategory.OST_StructuralFoundation))
                    is_vertical = (cat_val in [int(BuiltInCategory.OST_StructuralColumns), int(BuiltInCategory.OST_Walls)])
                    
                    if normal.Z > 0.9: continue 
                    if (is_foundation or is_vertical) and normal.Z < -0.9: continue
                    
                    try:
                        loops = face.GetEdgesAsCurveLoops()
                        if not loops: continue
                        
                        # Generate the Initial Panel
                        if solid_opts:
                            panel = GeometryCreationUtilities.CreateExtrusionGeometry(List[CurveLoop](loops), normal, formwork_thickness, solid_opts)
                        else:
                            panel = GeometryCreationUtilities.CreateExtrusionGeometry(List[CurveLoop](loops), normal, formwork_thickness)
                        
                        if not panel or panel.Volume == 0: continue
                        
                        # ==========================================
                        # AUTO-CUT LOGIC (Difference)
                        # ==========================================
                        for invader in invading_solids:
                            try:
                                # This cuts the intersecting structural element OUT OF the formwork panel
                                panel = BooleanOperationsUtils.ExecuteBooleanOperation(panel, invader, BooleanOperationsType.Difference)
                            except:
                                # Ignore if boolean fails due to micro-precision, keep the current panel
                                pass 
                        
                        if panel and panel.Volume > 0: 
                            element_formwork_panels.Add(panel)
                    except: pass
                
                # Create the final cut DirectShape
                if element_formwork_panels.Count > 0:
                    ds = DirectShape.CreateElement(doc, generic_model_cat)
                    ds.AppendShape(element_formwork_panels)
                    
                    param_comments = ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                    if param_comments: param_comments.Set("Formwork - {}".format(category_string))
                    
                    param_mark = ds.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                    if param_mark: param_mark.Set("HostID: {}".format(elem.Id.ToString()))
                        
                    total_generated += 1
                
                pb.update_progress(i, len(elements))

        if ui.generate_mto: 
            create_formwork_mto(doc)
            
    forms.alert("Successfully generated {} precision-cut formwork objects.\n\nFormwork is sliced per element AND automatically has holes where other structural elements intersect (Balok menembus Kolom, dll).".format(total_generated), title="Success")

if __name__ == '__main__':
    main()
