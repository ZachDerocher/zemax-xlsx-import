import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pandas.plotting import table

def render_mpl_table(data, col_width=6.0, row_height=0.625, font_size=10):
    """
    plots a pandas df via matplotlib, with some basic formatting
    """

    data_print=data.round(4)
    #data_print[data_print >= 1e10] = np.inf

    size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
    f, a = plt.subplots()
    a.axis('off')
    mpl_table = a.table(cellText=data_print.values, bbox=[0, 0, 1, 1], colLabels=data.columns)
    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)
    plt.show(block=True)


    return


def display_lde(TheSystem):
    # build a temporary dataframe, reading values from LDE, then print

    nsurf = TheSystem.LDE.NumberOfSurfaces
    lde = pd.DataFrame({
            'surftype': pd.Series(dtype='str'),
            'comment': pd.Series(dtype='str'),
            'radius': pd.Series(dtype='float'),
            'thickness': pd.Series(dtype='float'),
            'material': pd.Series(dtype='str'),
            'clearSD': pd.Series(dtype='float'),
            'chip': pd.Series(dtype='float'),
            'mechSD': pd.Series(dtype='float'),
            'conic': pd.Series(dtype='float'),
            }, index=np.arange(nsurf))

    # print the values
    for s in range(0, nsurf):
        lde.loc[s,'surftype'] = TheSystem.LDE.GetSurfaceAt(s).Type.ToString()
        lde.loc[s,'comment'] = TheSystem.LDE.GetSurfaceAt(s).Comment
        lde.loc[s,'radius'] = TheSystem.LDE.GetSurfaceAt(s).Radius
        lde.loc[s,'thickness'] = TheSystem.LDE.GetSurfaceAt(s).Thickness
        lde.loc[s,'material'] = TheSystem.LDE.GetSurfaceAt(s).Material
        lde.loc[s,'clearSD'] = TheSystem.LDE.GetSurfaceAt(s).SemiDiameter
        lde.loc[s,'chip'] = TheSystem.LDE.GetSurfaceAt(s).ChipZone
        lde.loc[s,'mechSD'] = TheSystem.LDE.GetSurfaceAt(s).MechanicalSemiDiameter
        lde.loc[s,'conic'] = TheSystem.LDE.GetSurfaceAt(s).Conic
        
    # display the LDE table
    print(lde)

    # better yet, make a table with mpl
    # this is giving a weird error, but works...
    #render_mpl_table(lde, col_width=1.0, font_size=10)


def set_system_units(TheSystem, PatentData, ZOSAPI):
    unit = PatentData['META']['lens_unit'][0].lower()
    if (unit=='meters') | (unit=='meter') | (unit=='m'):
        TheSystem.SystemData.Units.LensUnits = ZOSAPI.SystemData.ZemaxSystemUnits.Meters
    elif (unit=='inches') | (unit=='inch') | (unit=='in'):
        TheSystem.SystemData.Units.LensUnits = ZOSAPI.SystemData.ZemaxSystemUnits.Inches
    elif (unit=='centimeters') | (unit=='centimeter') | (unit=='cm'):
        TheSystem.SystemData.Units.LensUnits = ZOSAPI.SystemData.ZemaxSystemUnits.Centimeters
    else:
        TheSystem.SystemData.Units.LensUnits = ZOSAPI.SystemData.ZemaxSystemUnits.Millimeters


def insert_surfaces(TheSystem, PatentData):
    # add required number of rows to LDE
    for s in range(0, len(PatentData['SURF']['surf_num']) - 2):
        TheSystem.LDE.InsertNewSurfaceAt(2)

    # set the stop, based on '_STO' substring
    surf_data = PatentData['SURF']
    stop_surf = surf_data[surf_data['surf_num'].str.lower().str.contains('_sto')].index[0].item()
    if not stop_surf:
        print("\nERROR: no stop surface found")
        print("ensure there is a '_STO' substring in the surface number data\n\n")
    TheSystem.LDE.GetSurfaceAt(stop_surf).IsStop = True


def set_surface_types(TheSystem, PatentData, ZOSAPI):
    if 'ASPH' not in PatentData.keys():
        return

    # note: we only allow extended asphere for now
    TmpSurf = TheSystem.LDE.GetSurfaceAt(1)
    Asphere = TmpSurf.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.ExtendedOddAsphere)
    # can add other types here...

    for s in range(0, len(PatentData['ASPH']['surf_num'])):
        ThisSurf = TheSystem.LDE.GetSurfaceAt(PatentData['ASPH']['surf_num'][s])
        ThisSurf.ChangeType(Asphere)


def set_glass_catalogs(TheSystem, PatentData):
    # add all the required glass catalogs into the system explorer
    materials = PatentData['SURF']['nd'].dropna()
    catalogs = PatentData['SURF']['vd'].dropna()
    catalogs_to_use = []
    for i in materials.index.values:
        material_name = materials[i]
        try: 
            catalog_name = catalogs[i]
        except:
            print("error: each SURF 'nd' value must have corresponding 'vd' value")
        if type(material_name) != str:
            continue
        if catalog_name not in catalogs_to_use:
            catalogs_to_use.append(catalog_name)
        # also check here if the glass name exists in the catalog
        if material_name not in TheSystem.SystemData.MaterialCatalogs.GetMaterialsInCatalog(catalog_name):
            print(f"error: material {material_name} not found in catalog {catalog_name}\n")

    # add all the catalogs to the System Explorer
    available_catalogs = TheSystem.SystemData.MaterialCatalogs.GetAvailableCatalogs()
    default_catalogs = TheSystem.SystemData.MaterialCatalogs.GetCatalogsInUse()
    for c in catalogs_to_use:
        if c not in available_catalogs:
            print(f"error: glass catalog {c} not found in AGF file\n")
        else:
            if c not in default_catalogs:
                TheSystem.SystemData.MaterialCatalogs.AddCatalog(c)
    # remove the unused default catalog(s)
    for c_default in default_catalogs:
        if c_default not in catalogs_to_use:
            TheSystem.SystemData.MaterialCatalogs.RemoveCatalog(c_default)


def set_surface_data(TheSystem, PatentData, ZOSAPI):
    # set all the surface data values

    # set all standard surface data
    for s in range(0, len(PatentData['SURF']['surf_num'])):
        this_surf = TheSystem.LDE.GetSurfaceAt(s)

        # address by key name (alternatively, could use column#...)
        # apply comment
        this_surf.Comment = PatentData['SURF']['surf_num'][s]
        
        # apply radius of curvature
        if isinstance(PatentData['SURF']['r'][s], str):
            # allow string (i.e. "INF") for radius
            this_surf.Radius = np.inf
        else:
            this_surf.Radius = PatentData['SURF']['r'][s]
        
        # apply thickness
        if isinstance(PatentData['SURF']['d'][s], str):
            # allow string (i.e. "INF") for thickness
            this_surf.Thickness = np.inf
        else:
            this_surf.Thickness = PatentData['SURF']['d'][s]
        
        # apply material (if necessary)
        if type(PatentData['SURF']['nd'][s]) == str:
            # set material name (we already added the glass catalogs to System Explorer)
            this_surf.Material = PatentData['SURF']['nd'][s]
        else:
            # it is a numeric value (either nd, or nan i.e. no value)
            if not np.isnan(PatentData['SURF']['nd'][s]):
                # set material solve by Nd/Vd
                material_solve = this_surf.MaterialCell.CreateSolveType(ZOSAPI.Editors.SolveType.MaterialModel)._S_MaterialModel
                material_solve.IndexNd = PatentData['SURF']['nd'][s]  
                material_solve.AbbeVd = PatentData['SURF']['vd'][s]
                this_surf.MaterialCell.SetSolveData(material_solve)
            else:
                # no material value; do nothing
                pass
        
        # apply semi-diameter (if necessary)
        if not np.isnan(PatentData['SURF']['cir'][s]):
            this_surf.SemiDiameter = PatentData['SURF']['cir'][s]
            this_surf.MechanicalSemiDiameter = PatentData['SURF']['cir'][s]

        
    if 'ASPH' not in PatentData.keys():
        return
    # set all aspheric data
    asphere_data = PatentData['ASPH']
    keys = asphere_data.keys()
    coeff_indices = []
    for k in range(2, len(keys)):
        this_coeff_index = int(keys[k].split("_")[1]) # parse out coefficient index
        coeff_indices.append(this_coeff_index)
    for s in range(0, len(asphere_data['surf_num'])):
        ThisSurf = TheSystem.LDE.GetSurfaceAt(asphere_data['surf_num'][s])
        # set conic constant
        ThisSurf.Conic = asphere_data['ka'][s]
        # set asphere meta params
        num_asphere_terms = len(asphere_data.columns)-2 # don't count name, and conic
        ThisSurf.GetCellAt(24).IntegerValue = max(coeff_indices) # cell 24 is par(13), the max # of asphere terms
        ThisSurf.GetCellAt(25).DoubleValue = 1.0 # cell 25 is par(14), the normalization radius
        # set asphere coefficients
        for i in range(0, len(coeff_indices)):
            cell_index = 25 + coeff_indices[i]
            asphere_value = asphere_data[keys[i+2]][s]
            ThisSurf.GetCellAt(cell_index).DoubleValue = float(asphere_value)


def set_wavelengths(TheSystem, PatentData):
    # edit the pre-existing default
    TheSystem.SystemData.Wavelengths.GetWavelength(1).Wavelength = 0.001*PatentData['WAVE']['wavelength_nm'][0]
    TheSystem.SystemData.Wavelengths.GetWavelength(1).Weight = PatentData['WAVE']['weight'][0]

    # add the other wavelengths and weights
    for w in range(1, len(PatentData['WAVE']['wave_num'])):
        this_wavelength = 0.001*PatentData['WAVE']['wavelength_nm'][w]
        this_weight = PatentData['WAVE']['weight'][w]
        TheSystem.SystemData.Wavelengths.AddWavelength(this_wavelength, this_weight)
    
    # set the primary wavelength
    wave_data = PatentData['WAVE']
    pwave = wave_data[wave_data['wave_num'].str.lower().str.contains('_c')].index[0].item()
    if not pwave:
        print("\nERROR: no primary wavelength found")
        print("ensure there is a '_c' substring in the wave_num data\n\n")
    # the indexing if off by 1
    TheSystem.SystemData.Wavelengths.GetWavelength(pwave+1).MakePrimary()


def set_system_data(TheSystem, PatentData, ZOSAPI):
    # set all the system data, based on the first column of the CONF group
    # interpretation of CONF data:
    # d_n : (ignored in this function; see "set_mce()")
    # fno : system f/#
    # y_n : field height, for field number n
    
    # turn on ray aiming
    TheSystem.SystemData.RayAiming.RayAiming = ZOSAPI.SystemData.RayAimingMethod.Real

    # break out each relevant data type
    conf_data = PatentData['CONF']
    fno_data = conf_data[conf_data['name'].str.lower().str.contains('fno')].reset_index(drop=True)
    field_data = conf_data[conf_data['name'].str.lower().str.contains('y_')].reset_index(drop=True)

    # set system aperture
    # assume image space f/# (could read aperture type from excel...)
    TheSystem.SystemData.Aperture.ApertureType = ZOSAPI.SystemData.ZemaxApertureType.ImageSpaceFNum
    TheSystem.SystemData.Aperture.ApertureValue = fno_data['config_1'].item()

    # set the fields
    # assume real image height (could read field type from excel...)
    TheSystem.SystemData.Fields.SetFieldType(ZOSAPI.SystemData.FieldType.RealImageHeight)
    TheSystem.SystemData.Fields.GetField(1).Y = field_data['config_1'][0]
    for f in range(1, len(field_data['name'])):
        this_x = 0
        this_y = field_data['config_1'][f]
        this_w = 1
        TheSystem.SystemData.Fields.AddField(this_x, this_y, this_w)


def set_mce_data(TheSystem, PatentData, ZOSAPI):
    # set all the multi-config data, based on the first column of the CONF group
    # interpretation of CONF data:
    # d_n : thickness value, of surface n
    # fno : system f/#
    # y_n : field height, for field number n
    
    # first, check if there are multi-configs to add
    if len(PatentData['CONF'].columns) <= 2:
        return
    else:
        # add the necessary extra configs (beyond 1)
        for c in range(2, len(PatentData['CONF'].columns)):
            TheSystem.MCE.AddConfiguration(False)

    # break out each relevant data type
    conf_data = PatentData['CONF']
    fno_data = conf_data[conf_data['name'].str.lower().str.contains('fno')].reset_index(drop=True)
    field_data = conf_data[conf_data['name'].str.lower().str.contains('y_')].reset_index(drop=True)
    thickness_data = conf_data[conf_data['name'].str.lower().str.contains('d_')].reset_index(drop=True)
    keys = conf_data.keys()

    # fno operand (assume only 1...)
    op_fno = TheSystem.MCE.AddOperand()
    op_fno.ChangeType(ZOSAPI.Editors.MCE.MultiConfigOperandType.APER)
    # set aperture value for each config 
    for c in range(1, len(PatentData['CONF'].columns)):
        op_fno.GetOperandCell(c).DoubleValue = fno_data[keys[c]][0]

    # field operands
    for f in range(0, len(field_data)):
        # add operand
        op_field = TheSystem.MCE.AddOperand()
        op_field.ChangeType(ZOSAPI.Editors.MCE.MultiConfigOperandType.YFIE)
        # set the field number
        field_num = int(field_data['name'].iloc[f].split("_")[1]) # parse the field number
        op_field.Param1 = field_num-1 # Zemax is expecting the index of the field drop-down list (so field#1 is entered as '0', etc.)
        # set field Y value for each config 
        for c in range(1, len(PatentData['CONF'].columns)):
            this_field = field_data[keys[c]][f]
            op_field.GetOperandCell(c).DoubleValue = this_field

    # thickness operands
    for d in range(0, len(thickness_data)):
        # add operand
        op_thickness = TheSystem.MCE.AddOperand()
        op_thickness.ChangeType(ZOSAPI.Editors.MCE.MultiConfigOperandType.THIC)
        # set the surface number
        surf_num = int(thickness_data['name'].iloc[d].split("_")[1]) # parse the surface number
        op_thickness.Param1 = surf_num
        # set surface thickness value for each config 
        for c in range(1, len(PatentData['CONF'].columns)):
            this_thickness = thickness_data[keys[c]][d]
            if isinstance(this_thickness, str):
                # allow string (i.e. "INF") for thickness
                op_thickness.GetOperandCell(c).DoubleValue = 1e10
            else:
                op_thickness.GetOperandCell(c).DoubleValue = this_thickness


def write_patent_data_to_zemax(PatentData, zos, out_fn):
    # writes data into the lens data editor of a new optical system

    # system/variable prep
    import ZOSAPI
    TheApplication = zos.TheApplication
    TheSystem = TheApplication.PrimarySystem
    TheSystem.New(False)

    # META: set the system units
    set_system_units(TheSystem, PatentData, ZOSAPI)

    # SURF: set the number of surfaces, and STOP
    insert_surfaces(TheSystem, PatentData)

    # SURF: set the surface types
    set_surface_types(TheSystem, PatentData, ZOSAPI)
    
    # SURF: add material catalogs to System Explorer
    set_glass_catalogs(TheSystem, PatentData)

    # SURF: set each surface value in the LDE
    set_surface_data(TheSystem, PatentData, ZOSAPI)

    # WAVE: set the system wavelengths
    set_wavelengths(TheSystem, PatentData)

    # CONF: set the system data, using values from first column of CONF block
    set_system_data(TheSystem, PatentData, ZOSAPI)

    # CONF: set multi-config data, if it is provided
    set_mce_data(TheSystem, PatentData, ZOSAPI)

    # display the LDE
    display_lde(TheSystem)

    # save the system
    TheSystem.SaveAs(out_fn)