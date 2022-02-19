#Author-Dmitry Lisitsyn
#Description-

from html import entities
import sys, subprocess
import os
import inspect
import adsk.core, adsk.fusion, adsk.cam, traceback
import math
import tkinter as tk
from tkinter import  filedialog
import json
_app = adsk.core.Application.get()
_ui = _app.userInterface

script_path = os.path.abspath(inspect.getfile(inspect.currentframe()))
script_name = os.path.splitext(os.path.basename(script_path))[0]
script_dir = os.path.dirname(script_path)

sys.path.append(script_dir + '\modules')

try:
    from .modules.docx import docx   
    from .modules.fpdf2.fpdf import FPDF
except:
    _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'fpdf2'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'python-docx'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'Pillow'])
del sys.path[-1]

# Globals
_units = ''
commandId = 'WormGear'
commandName = 'Worm Gear Generator'
commandDescription = 'Worm Gear Generator'
commandToolTip = "Worm Gear Generator"
toolbarPanels = ["NewPanel"]
_Aw_tab = adsk.core.TextBoxCommandInput.cast(None)
_Module_tab = adsk.core.TextBoxCommandInput.cast(None)
Eps_tab = adsk.core.TextBoxCommandInput.cast(None)
Module_ = adsk.core.DropDownCommandInput.cast(None)
selectCreateWorm = adsk.core.DropDownCommandInput.cast(None)
selectCreateGear = adsk.core.DropDownCommandInput.cast(None)
Angle_prof_ = adsk.core.DropDownCommandInput.cast(None)
Angle_teeth_ = adsk.core.ValueCommandInput.cast(None)
Koef_smesh_ = adsk.core.ValueCommandInput.cast(None)
Num_of_vit_worm = adsk.core.ValueCommandInput.cast(None)
Peredat_Changed = adsk.core.ValueCommandInput.cast(None)
Length_wormNarez = adsk.core.ValueCommandInput.cast(None)
KolOborotov_worm = adsk.core.IntegerSpinnerCommandInput.cast(None)
Av_diam_worm = adsk.core.ValueCommandInput.cast(None)
Width_WG = adsk.core.ValueCommandInput.cast(None)
Koef_Diam_worm = adsk.core.ValueCommandInput.cast(None)
Teeth_Num_Gear = adsk.core.ValueCommandInput.cast(None)
Koef_smesh_Gear = adsk.core.ValueCommandInput.cast(None)
_naruz_Diam_tab = adsk.core.TextBoxCommandInput.cast(None)
_Alpha_tab =adsk.core.TextBoxCommandInput.cast(None)
_Df_tab = adsk.core.TextBoxCommandInput.cast(None)
_Dsr_tab = adsk.core.TextBoxCommandInput.cast(None)
_Da_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
_D_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
_Df_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
_xmin_tab = adsk.core.TextBoxCommandInput.cast(None)
Moment_ = adsk.core.ValueCommandInput.cast(None)
Power_ = adsk.core.ValueCommandInput.cast(None)
Velocity_ = adsk.core.ValueCommandInput.cast(None)
Moment_WG = adsk.core.ValueCommandInput.cast(None)
Power_WG = adsk.core.ValueCommandInput.cast(None)
Velocity_WG = adsk.core.ValueCommandInput.cast(None)
Fw_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
Fd_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
Fa_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
Ft_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
Fa_worm_tab = adsk.core.TextBoxCommandInput.cast(None)
Ft_worm_tab = adsk.core.TextBoxCommandInput.cast(None)
Vk_tab = adsk.core.TextBoxCommandInput.cast(None)
Fn_tab = adsk.core.TextBoxCommandInput.cast(None)
Frad_tab = adsk.core.TextBoxCommandInput.cast(None)
Peredat_= adsk.core.DropDownCommandInput.cast(None)
KPD = adsk.core.ValueCommandInput.cast(None)
hole_diameter = adsk.core.ValueCommandInput.cast(None)
Kw = adsk.core.ValueCommandInput.cast(None)
Elastic = adsk.core.ValueCommandInput.cast(None)
Puasson = adsk.core.ValueCommandInput.cast(None)
Kmat = adsk.core.ValueCommandInput.cast(None)
y_Luis = adsk.core.ValueCommandInput.cast(None)
Sn = adsk.core.ValueCommandInput.cast(None)
Fs_WG_tab = adsk.core.TextBoxCommandInput.cast(None)
radio_CountType = adsk.core.RadioButtonGroupCommandInput.cast(None)
radio_WormSize = adsk.core.RadioButtonGroupCommandInput.cast(None)
radioButtonS = adsk.core.RadioButtonGroupCommandInput.cast(None)
Kpd_Check = adsk.core.BoolValueCommandInput.cast(None)
hole_Check = adsk.core.BoolValueCommandInput.cast(None)
buttonRowInput = adsk.core.ButtonRowCommandInput.cast(None)
buttonSaveLoad = adsk.core.ButtonRowCommandInput.cast(None)
buttonimportParams = adsk.core.ButtonRowCommandInput.cast(None)

_handlers = []
materialsMap = {}
tbPanel = None
Peredat = 0
ModuleExp = 0
NumVitkovExp = 0
PressureAngleExp = 0
DiamWormExp = 0
LengthWormExp = 0
materialWorm = None 

def run(context):
    try:
        global _app, _ui, commandId, commandName, commandDescription

        workSpace = _ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workSpace.toolbarPanels
        # faceSel = _ui.selectEntity('Select a face.', 'Faces')
        # if faceSel:
        #     selectedFace = adsk.fusion.BRepFace.cast(faceSel.entity)

        #     # Find the index of this face within the body.            
        #     faceIndex = -1
        #     body = selectedFace.body
        #     faceCount = 0
        #     for face in body.faces:
        #         if face == selectedFace:
        #             faceIndex = faceCount
        #             b reak
                
        #         faceCount += 1

        #     _ui.messageBox('The selected face is index ' + str(faceCount)) 
        global tbPanel
        tbPanel = tbPanels.itemById('NewPanel')
        if tbPanel:
            tbPanel.deleteMe()
        tbPanel = tbPanels.add('NewItem', 'Worm Gear Generator', 'SelectPanel', False)
        cmdDef = _ui.commandDefinitions.itemById('panel')

        if cmdDef:
            cmdDef.deleteMe()
        message = '<div>Генератор черячной передачи предназначен для автоматического расчета параметров червячной передачи и ее моделирования на основании введенных пользователем данных.\n<br> Результаты расчета меняются динамически в зависимости от изменяемых значений.\n<br> Пользователь может отдельно выбрать компонент для 3D моделирования.</div>'
        cmdDef = _ui.commandDefinitions.addButtonDefinition('panel', 'Worm Gear Generator', message,'Resources/Icon')
        cmdDef.toolClipFilename = 'resources/WormGear/WormGearTooltip.png'
        cmdDefcontrol=tbPanel.controls.addCommand(cmdDef)
        
        cmdDefcontrol.isPromotedByDefault=True
        cmdDefcontrol.isPromoted=True

        onCommandCreated = GearCommandCreatedHandler()
        cmdDef.commandCreated.add(onCommandCreated)
        _handlers.append(onCommandCreated)
    except:
        if _ui:
            _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    try:
        global _app, _ui
        if tbPanel:
            tbPanel.deleteMe()
        _ui.messageBox('Addin has been stopped')
    except:
        if _ui:
            _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# class GearCommandDestroyHandler(adsk.core.CommandEventHandler):
#     def __init__(self):
#         super().__init__()
#     def notify(self, args):
#         try:
#             eventArgs = adsk.core.CommandEventArgs.cast(args)

#             adsk.terminate()
#         except:
#             if _ui:
#                 _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def getCommandInputValue(commandInput, unitType):
    try:
        valCommandInput = adsk.core.ValueCommandInput.cast(commandInput)
        if not valCommandInput:
            return (False, 0)

        # Verify that the expression is valid.
        des = adsk.fusion.Design.cast(_app.activeProduct)
        unitsMgr = des.unitsManager
        
        if unitsMgr.isValidExpression(valCommandInput.expression, unitType):
            value = unitsMgr.evaluateExpression(valCommandInput.expression, unitType)
            return (True, value)
        else:
            return (False, 0)
    except:
        if _ui:
            _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Event handler for the commandCreated event.
class GearCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            des = adsk.fusion.Design.cast(_app.activeProduct)
            eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)

            cmd = eventArgs.command
            cmd.isExecutedWhenPreEmpted = False
            inputs = cmd.commandInputs
            
            global Kpd_Check,buttonRowInput, hole_diameter, hole_Check, buttonimportParams, buttonSaveLoad, selectCreateWorm, selectCreateGear, radio_WormSize, Puasson, Kmat, Peredat_Changed, KolOborotov_worm, radio_CountType, Moment_WG, Velocity_WG, Power_WG, Power_, radioButtonS, Sn, Fs_WG_tab, y_Luis, Kw, Elastic, KPD, Velocity_, Peredat_, Moment_, Fw_WG_tab, Fd_WG_tab, Fa_WG_tab, Ft_WG_tab, Fa_worm_tab, Ft_worm_tab, Vk_tab, Fn_tab, Frad_tab, Eps_tab, Width_WG, Angle_prof_, Angle_teeth_, _xmin_tab, _Df_WG_tab, _D_WG_tab, _Da_WG_tab, _Dsr_tab, _Df_tab, _Alpha_tab, _naruz_Diam_tab, Koef_smesh_Gear, Teeth_Num_Gear, Num_of_vit_worm, Length_wormNarez, Av_diam_worm, Koef_Diam_worm, commandId, _Aw_tab, Module_, _Module_tab
           
            # ВКЛАДКА МОДЕЛЬ
            # TAB МОДЕЛЬ
            tabCmdInput1 = inputs.addTabCommandInput(commandId + '_tab_1', 'Модель' , "resources/tabModel")
            tab1ChildInputs = tabCmdInput1.children

            # ОБЩИЕ
            groupCmdInput_General = tab1ChildInputs.addGroupCommandInput(commandId + '_groupGeneral', 'Общее')
            groupCmdInput_General.isExpanded = False
            groupChildInputs_General = groupCmdInput_General.children

            buttonSaveLoad = groupChildInputs_General.addButtonRowCommandInput('buttonSaveLoad', 'Данные', True)
            buttonSaveLoad.listItems.add('Сохранить введенные параметры в файл', False, 'resources/Menu_Save')
            buttonSaveLoad.listItems.add('Загрузить параметры из файла', False, 'resources/Menu_Load')

           
            radio_CountType = groupChildInputs_General.addRadioButtonGroupCommandInput('Model',
                                                                                     'Исходный параметр')
            radioButtonItems = radio_CountType.listItems
            radioButtonItems.add("Передаточное отношение", False)
            radioButtonItems.add("Количество зубьев", True)
            radio_CountType.tooltip = "Описание для выбора исходного параметра расчета"


            radio_WormSize = groupChildInputs_General.addRadioButtonGroupCommandInput('Model',
                                                                                     'Размер червяка')
            radioButtonItems = radio_WormSize.listItems
            radioButtonItems.add("Коэффициент диаметра", False)
            radioButtonItems.add("Угол наклона зуба", False)
            radioButtonItems.add("Средний диаметр", True)
            radio_WormSize.tooltip = "Описание для выбора исходного параметра расчета размера червяка"

            peredat = '20.0000'
            peredatAttrib = des.attributes.itemByName('WormGear', 'peredatNumber')
            if peredatAttrib:
                peredat = peredatAttrib.value

            Peredat_ =  groupChildInputs_General.addDropDownCommandInput('Model',
                                                                              'Требуемое передаточное отношение',
                                                               adsk.core.DropDownStyles.TextListDropDownStyle)
            dropdownItems = Peredat_.listItems
            Items = ['5.6000', '6.3000','7.1000',
            '8.0000','9.0000','10.0000','11.2000','12.5000',
            '14.0000','16.0000','18.0000','20.0000','22.4000',
            '25.0000','28.0000','31.5000','40.0000','45.0000',
            '50.0000','56.0000','63.0000','71.0000','80.0000','90.0000','100.0000']
            for i in Items:
                if(peredat == i):
                    dropdownItems.add(i, True, '')
                else:
                    dropdownItems.add(i, False, '')
           
            Peredat_.tooltip = "Передаточное отношение"
            Peredat_.tooltipDescription = "Передаточное число червячной передачи равно отношению числа зубьев червячного колеса к числу витков червяка.\n Передаточное число показывает, насколько данная пара зацепления в принципе может изменить крутящий момент в ту или иную сторону, линейное соотношение диаметров зубчатых колёс"
            
            Peredat_Changed =  groupChildInputs_General.addValueInput('Model', 'Требуемое передаточное отношение', '',
                                                adsk.core.ValueInput.createByReal(0.0))
            Peredat_Changed.tooltip = "Передаточное отношение"
            Peredat_Changed.tooltipDescription = "Передаточное число червячной передачи равно отношению числа зубьев червячного колеса к числу витков червяка\n Передаточное число показывает, насколько данная пара зацепления в принципе может изменить крутящий момент в ту или иную сторону, линейное соотношение диаметров зубчатых колёс"

            temp_Peredat = (Peredat_.selectedItem.name).split(' ')
            temp_Peredat = float(temp_Peredat[0])

            module = '4.000 mm'
            moduleAttrib = des.attributes.itemByName('WormGear', 'module')
            if moduleAttrib:
                module = moduleAttrib.value
            
            Module_ = groupChildInputs_General.addDropDownCommandInput('Model', 'Модуль, [мм]',
                                                                              adsk.core.DropDownStyles.TextListDropDownStyle)
            dropdownItems = Module_.listItems
            Items = ['1.000 mm', '1.150 mm','1.250 mm',
            '1.400 mm','1.600 mm','1.800 mm','2.000 mm','2.240 mm',
            '2.500 mm','2.800 mm','3.150 mm','3.550 mm','4.000 mm',
            '4.500 mm','5.000 mm','5.600 mm','6.300 mm','7.100 mm',
            '8.000 mm','9.000 mm','10.000 mm','11.200 mm','12.500 mm','14.000 mm','16.000 mm', '18.000 mm',
            '20.000 mm', '22.400 mm', '25.000 mm']
            for i in Items:
                if(module == i):
                    dropdownItems.add(i, True, '')
                else:
                    dropdownItems.add(i, False, '')

            Module_.tooltip = "Модуль"
            Module_.tooltipDescription = "Модуль – число миллиметров диаметра делительной окружности, приходящееся на один зуб. Самый главный параметр, стандартизирован, определяется из прочностного расчёта зубчатых передач. Чем больше нагружена передача, тем выше значение модуля. Через него выражаются все остальные параметры"
            temp_Module = (Module_.selectedItem.name).split(' ')
            temp_Module = float(temp_Module[0])


            pressureAngle = '20.0000 deg'
            presAngleAttrib = des.attributes.itemByName('WormGear', 'pressureAngle')
            if presAngleAttrib:
                pressureAngle = presAngleAttrib.value
            Items = ['14.5000 deg','17.5000 deg', '20.0000 deg', '22.5000 deg', '25.0000 deg', '30.0000 deg']

            Angle_prof_ = groupChildInputs_General.addDropDownCommandInput('Model',
                                                                              'Угол профиля, [град]',
                                                                              adsk.core.DropDownStyles.TextListDropDownStyle)
            dropdownItems = Angle_prof_.listItems
            for i in Items:
                if(pressureAngle == i):
                    dropdownItems.add(i, True, '')
                else:
                    dropdownItems.add(i, False, '')
            Angle_prof_.tooltip = "Угол профиля"
            Angle_prof_.tooltipDescription = "Острый угол между касательной к профилю в данной точке и радиусом - вектором, проведенным в данную точку из центра"
            Angle_prof_.toolClipFilename = 'resources/WormGear/pressureAngle.png'

           
            temp_Angle_prof = (Angle_prof_.selectedItem.name).split(' ')
            temp_Angle_prof = float(temp_Angle_prof[0])

            Angle_teeth_ = groupChildInputs_General.addValueInput('Model', 'Угол наклона зуба, [град]', '',
                                                   adsk.core.ValueInput.createByReal(0.0))
            Angle_teeth_.tooltip = "Угол наклона зуба"
            # Червяк
            groupCmdInput_Worm = tab1ChildInputs.addGroupCommandInput('Model', 'Червяк')
            groupCmdInput_Worm.isExpanded = False
            groupChildInputs_Worm = groupCmdInput_Worm.children

            buildWorm = 'Построить 3D модель'
            buildWormAttrib = des.attributes.itemByName('Worm', 'BuildWorm')
            if buildWormAttrib:
                buildWorm = buildWormAttrib.value
            selectCreateWorm = groupChildInputs_Worm.addDropDownCommandInput('Model', 'Компонент',
                                                                           adsk.core.DropDownStyles.TextListDropDownStyle)
            dropdownItems = selectCreateWorm.listItems
            if(buildWorm == 'Построить 3D модель'):
                dropdownItems.add('Построить 3D модель', True, '')
            else:
                dropdownItems.add('Построить 3D модель', False, '')
            if(buildWorm == 'Без модели'):
                dropdownItems.add('Без модели', True, '')
            else:
                dropdownItems.add('Без модели', False, '')
            
            selectCreateWorm.tooltip = "Построение компонента"

            Num_of_vit_worm =  groupChildInputs_Worm.addIntegerSpinnerCommandInput('Model', 'Количество витков', 1, 20, 1, 4)

            NumVitkovAttrib = des.attributes.itemByName('Worm', 'NumVitkov')
            if NumVitkovAttrib:
                Num_of_vit_worm.value = int(NumVitkovAttrib.value)


            Num_of_vit_worm.tooltip = "Количество витков червяка"
            Num_of_vit_worm.tooltipDescription = "Параметр задает количество витков червяка. От данного параметра зависит передаточное отношение червячной передачи и угол наклона зубьев"
            Num_of_vit_worm.toolClipFilename = 'resources/Worm/numVitkov.png'

            KolOborotov_worm = groupChildInputs_Worm.addIntegerSpinnerCommandInput('Model', 'Количество оборотов витков ', 1, 50, 1, 10)

            KolOborotovAttrib = des.attributes.itemByName('Worm', 'KolOborotov')
            if KolOborotovAttrib:
                KolOborotov_worm.value = int(KolOborotovAttrib.value)

            KolOborotov_worm.tooltip = "Количесто оборотов витков"
            KolOborotov_worm.tooltipDescription = "Параметр задает количество оборотов витков червяка. От данного параметра зависит длина нарезной части червяка"
            KolOborotov_worm.toolClipFilename = 'resources/Worm/numOborotov.png'

            Length_wormNarez = groupChildInputs_Worm.addValueInput('Model', 'Длина нарезной части червяка, [мм]', '',
                                                   adsk.core.ValueInput.createByReal(int(KolOborotov_worm.value)*temp_Module*3.14))
            Length_wormNarez.tooltip = "Длина нарезной части червяка"
            Length_wormNarez.tooltipDescription = "Длина нарезной части червяка зависит от модуля червячной передачи и количества оборотов витков червяка"
            Length_wormNarez.toolClipFilename = 'resources/Worm/lengthNarez.png'
            Length_wormNarez.isEnabled = False
            # Length_worm = groupChildInputs_Worm.addIntegerSpinnerCommandInput('Model', 'Длина червяка, [мм]', 1, 1000, 1, 140)
                                                
            # Length_worm.tooltip = "Длина червяка"
            # Length_worm.tooltipDescription = "Параметр задает длину проектируемого червяка"
            # Length_worm.toolClipFilename = 'resources/WormGear/quest.png'
            
            Koef_Diam_worm = groupChildInputs_Worm.addValueInput('Model', 'Коэффициент диаметра', '',
                                                adsk.core.ValueInput.createByReal(10))
            Koef_Diam_worm.tooltip = "Коэффициент диаметра"
            Koef_Diam_worm.tooltipDescription = "Коэффициент диаметра червяка q является отношением делительного диаметра червяка к его расчетному модулю"

            Av_diam_worm = groupChildInputs_Worm.addValueInput('Model', 'Средний диаметр, [мм]', '',
                                                adsk.core.ValueInput.createByReal(temp_Module * float(Koef_Diam_worm.value)))

            AvDiamAttrib = des.attributes.itemByName('Worm', 'AverageDiam')
            if AvDiamAttrib:
                Av_diam_worm.value = float(AvDiamAttrib.value)

            Av_diam_worm.tooltip = "Средний диаметр"
            Av_diam_worm.tooltipDescription = "Диаметр средней концентрической окружности червяка"
            Av_diam_worm.toolClipFilename = 'resources/Worm/diamWorm.png'

            Angle_teeth_.value = math.atan(float(Num_of_vit_worm.value)/(float(Koef_Diam_worm.value)))*180/math.pi   

            # Червячная передача
            groupCmdInput_WormGear = tab1ChildInputs.addGroupCommandInput('Model',
                                                                          'Червячное Колесо')
            groupCmdInput_WormGear.isExpanded = False
            groupChildInputs_WormGear = groupCmdInput_WormGear.children

            buildGear = 'Построить 3D модель'
            selectGearAttrib = des.attributes.itemByName('Gear', 'BuildGear')
            if selectGearAttrib:
                buildGear = selectGearAttrib.value
            selectCreateGear = groupChildInputs_WormGear.addDropDownCommandInput('Model', 'Компонент',
                                                                               adsk.core.DropDownStyles.TextListDropDownStyle)
            dropdownItems = selectCreateGear.listItems
            if(buildGear == 'Построить 3D модель'):
                dropdownItems.add('Построить 3D модель', True, '')
            else:
                dropdownItems.add('Построить 3D модель', False, '')
            if(buildGear == 'Без модели'):
                dropdownItems.add('Без модели', True, '')
            else:
                dropdownItems.add('Без модели', False, '')
            selectCreateGear.tooltip = "Построение компонента"

            Teeth_Num_Gear = groupChildInputs_WormGear.addIntegerSpinnerCommandInput('Model', 'Количество зубьев', 1, 1000, 1, 60)
            numTeethAttrib = des.attributes.itemByName('Gear', 'numTeeth')
            if numTeethAttrib:
                Teeth_Num_Gear.value = int(numTeethAttrib.value)

            Teeth_Num_Gear.tooltip = "Количество зубьев"
            Teeth_Num_Gear.tooltipDescription = "Количество зубьев на червячном колесе. От данного параметра зависит передаточное число"
            Teeth_Num_Gear.toolClipFilename = 'resources/WormGear/teethGear.png'

            Width_WG = groupChildInputs_WormGear.addIntegerSpinnerCommandInput('Model', 'Ширина грани, [мм]', 1, 100, 1, 20)
            thicknessAttrib = des.attributes.itemByName('Gear', 'thickness')
            if thicknessAttrib:
                Width_WG.value = int(thicknessAttrib.value)

            Width_WG.tooltip = "Ширина грани червячного колеса"
            Width_WG.tooltipDescription = "Параметр задает ширину проектируемого червячного колеса"
            Width_WG.toolClipFilename = 'resources/WormGear/widthWheel.png'

            Koef_smesh_Gear = groupChildInputs_WormGear.addValueInput('Model', 'Коэффициент смещения', '',
                                                    adsk.core.ValueInput.createByReal(1.0))
            koefSmeshAttrib = des.attributes.itemByName('Gear', 'koefSmesh')
            if koefSmeshAttrib:
                Koef_smesh_Gear.value = float(koefSmeshAttrib.value)

            Koef_smesh_Gear.tooltip = "Коэффициент смещения"
            Koef_smesh_Gear.tooltipDescription = "Коэффициент и определяет геометрию колеса (с подрезанием или нет), положительное значение – перемещение червяка от колеса, отрицательное - перемещение червяка к колесу"
            
            naprVitkov = 'Левое'
            naprVitkovAttrib = des.attributes.itemByName('Gear', 'NaprVitkov')
            if naprVitkovAttrib:
                naprVitkov = naprVitkovAttrib.value
            buttonRowInput = groupChildInputs_WormGear.addButtonRowCommandInput('buttonRow', 'Направление зубьев', False)
            if(naprVitkov == 'Левое'):
                buttonRowInput.listItems.add('Левое', True, 'resources/Left')
            else:
                buttonRowInput.listItems.add('Левое', False, 'resources/Left')
            if(naprVitkov == 'Правое'):
                buttonRowInput.listItems.add('Правое', True, 'resources/Right')
            else:
                buttonRowInput.listItems.add('Правое', False, 'resources/Right')
            
            hole_Check = groupChildInputs_WormGear.addBoolValueInput('Model', 'Добавить отверстие, [мм]', True, '', False)
            hole_diameter = groupChildInputs_WormGear.addValueInput('Model', '', '',
                                                adsk.core.ValueInput.createByReal(0))
            
            holeDiamAttrib = des.attributes.itemByName('Gear', 'holediam')
            if holeDiamAttrib:
                hole_Check.value = True
                hole_diameter.value = float(holeDiamAttrib.value)
               

            groupCmdInput_FirstResults = tab1ChildInputs.addGroupCommandInput('Model', 'Результаты')
            groupCmdInput_FirstResults.isExpanded = False
            groupChildInputs_Results = groupCmdInput_FirstResults.children

            # Таблица Результата Первого расчета
            # General_Results
            temp_d1 = float(Koef_Diam_worm.value) * temp_Module
            temp_d2 = int(Teeth_Num_Gear.value) * temp_Module
            temp_dw1 = temp_d1 + 2 * float(Koef_smesh_Gear.value) * temp_Module
            temp_dw2 = temp_d2

            table = adsk.core.TableCommandInput.cast(
                groupChildInputs_Results.addTableCommandInput('table0', 'Inputs', 2, '1:1'))
            table.minimumVisibleRows = 3
            table.maximumVisibleRows = 4
            table.columnSpacing = 1
            table.rowSpacing = 1
            table.tablePresentationStyle = adsk.core.TablePresentationStyles.itemBorderTablePresentationStyle
            table.hasGrid = True
            # 1
            text = tab1ChildInputs.addStringValueInput('text0', 'aw',
                                                       'Межосевое расстояние (aw)')
            text.isReadOnly = True
            table.addCommandInput(text, 0, 0, False, False)
            _Aw_tab = tab1ChildInputs.addTextBoxCommandInput('textAw', 'textFr',
                                                         '', 1, True)
            _Aw_tab.isReadOnly = True
            _Aw_tab.text = str((temp_dw1+temp_dw2)/2) + ' мм'
            table.addCommandInput(_Aw_tab, 0, 1, False, False)
            # 2
            text = tab1ChildInputs.addStringValueInput('text0', '',
                                                       'Коэффициент осевого перекрытия (ε)')
            text.isReadOnly = True
            table.addCommandInput(text, 1, 0, False, False)

            Eps_tab = tab1ChildInputs.addTextBoxCommandInput('textEps', 'textEps',
                                                         '', 1, True)
            Eps_tab.isReadOnly = True
            table.addCommandInput(Eps_tab, 1, 1, False, False)
            # 3
            text = tab1ChildInputs.addStringValueInput('text0', 'm',
                                                       'Модуль (m)')
            text.isReadOnly = True
            table.addCommandInput(text, 2, 0, False, False)
            _Module_tab = tab1ChildInputs.addTextBoxCommandInput('textModule', 'textModule',
                                                         '', 1, True)
            _Module_tab.isReadOnly = True
            _Module_tab.text = str(temp_Module) + ' мм'
            table.addCommandInput(_Module_tab, 2, 1, False, False)
            # 4
            text = tab1ChildInputs.addStringValueInput('text0', 'α',
                                                       'Угол профиля (α)')
            text.isReadOnly = True
            table.addCommandInput(text, 3, 0, False, False)
            _Alpha_tab = tab1ChildInputs.addTextBoxCommandInput('textAlpha', 'Alpha_text',
                                                         '', 1, True)
            _Alpha_tab.isReadOnly = True
            _Alpha_tab.text = str(((math.atan((math.tan(temp_Angle_prof * math.pi/180) * math.cos( float(Angle_teeth_.value))))) *180/math.pi)) + ' deg'
            table.addCommandInput(_Alpha_tab, 3, 1, False, False)

            # Results_Worm
            temp_d1 = float(Koef_Diam_worm.value) * temp_Module
            temp_da1 = temp_d1 +  2 * temp_Module * 1

            groupChildInputs_Results.addTextBoxCommandInput(commandId + '_textBox', '',
                                                            'Червяк', 1, True)
            table = adsk.core.TableCommandInput.cast(
                groupChildInputs_Results.addTableCommandInput('table3', 'Inputs', 2, '1:1'))
            table.minimumVisibleRows = 3
            table.maximumVisibleRows = 3
            table.columnSpacing = 1
            table.rowSpacing = 1
            table.tablePresentationStyle = adsk.core.TablePresentationStyles.itemBorderTablePresentationStyle
            table.hasGrid = True
            # 1
            text = tab1ChildInputs.addStringValueInput('text0', 'da',
                                                       'Наружный диаметр (da)')
            text.isReadOnly = True
            table.addCommandInput(text, 0, 0, False, False)

            _naruz_Diam_tab = tab1ChildInputs.addTextBoxCommandInput('naruz_Diam', 'naruz_Diam',
                                                       '', 1, True)
            _naruz_Diam_tab.isReadOnly = True
            _naruz_Diam_tab.text = str(temp_da1) + ' мм'
            table.addCommandInput(_naruz_Diam_tab, 0, 1, False, False)
            # 2
            text = tab1ChildInputs.addStringValueInput('text0', 'd',
                                                       'Средний диаметр (d)')
            text.isReadOnly = True
            table.addCommandInput(text, 1, 0, False, False)

            _Dsr_tab = tab1ChildInputs.addTextBoxCommandInput('Diam_sr', 'Diam_sr',
                                                       '', 1 , True)
            _Dsr_tab.isReadOnly = True
            _Dsr_tab.text = str(Av_diam_worm.value) + ' мм'
            table.addCommandInput(_Dsr_tab, 1, 1, False, False)
            # 3
            text = tab1ChildInputs.addStringValueInput('text0', 'df',
                                                       'Диаметр впадин (df)')
            text.isReadOnly = True
            table.addCommandInput(text, 2, 0, False, False)

            _Df_tab = tab1ChildInputs.addTextBoxCommandInput('Df_tab', 'Df_tab',
                                                       '', 1, True)

            _Df_tab.text = str(temp_d1-2*float(temp_Module)*(1+0.2)) + ' мм'                                           
            _Df_tab.isReadOnly = True
            table.addCommandInput(_Df_tab, 2, 1, False, False)

            # Results_WormGear
            
            groupChildInputs_Results.addTextBoxCommandInput(commandId + '_textBox', '',
                                                            'Червячная передача', 1, True)
            table = adsk.core.TableCommandInput.cast(
                groupChildInputs_Results.addTableCommandInput('table2', 'Inputs', 2, '1:1'))
            table.minimumVisibleRows = 3
            table.maximumVisibleRows = 3
            table.columnSpacing = 1
            table.rowSpacing = 1
            table.tablePresentationStyle = adsk.core.TablePresentationStyles.itemBorderTablePresentationStyle
            table.hasGrid = True

            # 1
            text = tab1ChildInputs.addStringValueInput('text0', 'da',
                                                       'Наружный диаметр (da)')
            text.isReadOnly = True
            table.addCommandInput(text, 0, 0, False, False)

            _Da_WG_tab = tab1ChildInputs.addTextBoxCommandInput('naruz_diam_WG', 'naruz_diam_WG',
                                                       '', 1, True)
            _Da_WG_tab.isReadOnly = True
            _Da_WG_tab.text = str(temp_d2+2 * temp_Module * (1 + float(Koef_smesh_Gear.value))) + 'мм'
            table.addCommandInput(_Da_WG_tab, 0, 1, False, False)
            # 2
            text = tab1ChildInputs.addStringValueInput('text0', 'd',
                                                       'Средний диаметр (d)')
            text.isReadOnly = True
            table.addCommandInput(text, 1, 0, False, False)

            _D_WG_tab = tab1ChildInputs.addTextBoxCommandInput('Dsr_WG', 'Dsr_WG',
                                                       '', 1 , True)
            _D_WG_tab.isReadOnly = True
            _D_WG_tab.text = str(int(Teeth_Num_Gear.value)*temp_Module) + ' мм'
            table.addCommandInput(_D_WG_tab, 1, 1, False, False)
            # 3
            text = tab1ChildInputs.addStringValueInput('text0', 'df',
                                                       'Диаметр впадин (df)')
            text.isReadOnly = True
            table.addCommandInput(text, 2, 0, False, False)
            _Df_WG_tab = tab1ChildInputs.addTextBoxCommandInput('df_WG', 'df_WG',
                                                       '', 1, True)
            _Df_WG_tab.isReadOnly = True
            _Df_WG_tab.text = str(temp_d2 - 2 * temp_Module *(1 + 0.2 - float(Koef_smesh_Gear.value))) + ' мм'
            table.addCommandInput(_Df_WG_tab, 2, 1, False, False)

            buttonimportParams = groupChildInputs_Results.addButtonRowCommandInput('buttonimportParams', 'Экспорт параметров', True)
            buttonimportParams.listItems.add('Экспорт параметров в PDF', False, 'resources/toPDF')
            buttonimportParams.listItems.add('Экспорт параметров в Word', False, 'resources/toWord')
            # 4
            # text = tab1ChildInputs.addStringValueInput('text0', 'Xmin',
            #                                            'Мин. рекоменд. коэффициент (Xmin)')
            # text.isReadOnly = True
            # table.addCommandInput(text, 3, 0, False, False)
            # _xmin_tab = tab1ChildInputs.addTextBoxCommandInput('xmin_WG', 'xmin_WG',
            #                                            '', 1, True)
            # _xmin_tab.isReadOnly = True
            # table.addCommandInput(_xmin_tab, 3, 1, False, False)

            # ВКЛАДКА РАСЧЕТ
            
            tabCmdInput2 = inputs.addTabCommandInput(commandId + '_tab_2', 'Расчет', 'resources/tabCalc')
            tab2ChildInputs = tabCmdInput2.children

            groupCmdInput_Load = tab2ChildInputs.addGroupCommandInput(commandId + '_groupGLoad', 'Нагрузка')
            groupCmdInput_Load.isExpanded = False
            groupChildInputs_Load = groupCmdInput_Load.children

            # Create radio button group input
            radioButtonS = groupChildInputs_Load.addRadioButtonGroupCommandInput('Model',
                                                                                     'Ведущий элемент')
            radioButtonItems = radioButtonS.listItems
            radioButtonItems.add("Червячная передача", False)
            radioButtonItems.add("Червяк", True)
            

            Power_ = groupChildInputs_Load.addValueInput('Model', 'Мощность (кВт)', '',
                                                adsk.core.ValueInput.createByString('0.100'))
            Power_.tooltip = "Мощность"
            Power_.tooltipDescription = "Величина, характеризующая мгновенную скорость передачи энергии от одной физической системы к другой в процессе её использования и в общем случае определяемая через соотношение переданной энергии к времени передачи"

            Velocity_ =  groupChildInputs_Load.addValueInput('Model', 'Скорость (об/мин)', '',
                                                adsk.core.ValueInput.createByString('1000.0'))
            Velocity_.tooltip = "Скорость"
            Velocity_.tooltipDescription = "Количество оборотов, выполняемых червячной передачей"

            temp_Moment = (60000*float(Power_.value))/(2*math.pi*float(Velocity_.value))
            Moment_ = groupChildInputs_Load.addValueInput('Model', 'Крутящий момент (Нм)', '',
                                                adsk.core.ValueInput.createByReal(temp_Moment))
            Moment_.tooltip = "Крутящий момент"

            Power_WG = groupChildInputs_Load.addValueInput('Model', 'Мощность (кВт)', '',
                                                adsk.core.ValueInput.createByString('0.084'))
            Velocity_WG =  groupChildInputs_Load.addValueInput('Model', 'Скорость (об/мин)', '',
                                                adsk.core.ValueInput.createByString('50.0'))
            temp_Moment_WG = (60000*float(Power_WG.value))/(2*math.pi*float(Velocity_WG.value))
            Moment_WG = groupChildInputs_Load.addValueInput('Model', 'Крутящий момент (Нм)', '',
                                                adsk.core.ValueInput.createByReal(temp_Moment_WG))
                
            Kpd_Check = groupChildInputs_Load.addBoolValueInput('Model', 'КПД', True, '', False)
            



            v1 = (math.pi*temp_d1*float(Velocity_.value))/(60000)
            v2 = v1*float(Num_of_vit_worm.value)/float(Koef_Diam_worm.value) 
            tgy = math.atan(float(Num_of_vit_worm.value)/float(Koef_Diam_worm.value))*180/math.pi
            vk = v1/math.cos(tgy*math.pi/180)
            phiz = math.atan(0.02+(0.03)/(vk))*180/math.pi
            kpd_Val = (math.tan(tgy*math.pi/180))/(math.tan((tgy+phiz)*math.pi/180))
            KPD = groupChildInputs_Load.addValueInput('Model', '', '',
                                                adsk.core.ValueInput.createByReal(kpd_Val*0.96))
            Kpd_Check.value = True


            groupCmdInput_MaterialWorm = tab2ChildInputs.addGroupCommandInput(commandId + '_groupMaterial',
                                                                          'Материал Червяка')
            groupCmdInput_MaterialWorm.isExpanded = False
            groupChildInputs_MaterialWorm = groupCmdInput_MaterialWorm.children

            materialLibInputWorm = groupChildInputs_MaterialWorm.addDropDownCommandInput(commandId + '_materialLibWorm', 'Библиотека материалов', adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            listItemsWorm = materialLibInputWorm.listItems
            materialLibNamesWorm = getMaterialLibNames()
            for materialName in materialLibNamesWorm :
                listItemsWorm.add(materialName, False, '')
            listItemsWorm[0].isSelected = True
            materialListInputWorm = groupChildInputs_MaterialWorm.addDropDownCommandInput(commandId + '_materialListWorm', 'Материал', adsk.core.DropDownStyles.TextListDropDownStyle)
            materialsWorm = getMaterialsFromLib(materialLibNamesWorm[0], '')
            listItemsWorm = materialListInputWorm.listItems
            for materialName in materialsWorm :
                listItemsWorm.add(materialName, False, '')
            listItemsWorm[0].isSelected = True
            filter_materialWorm = groupChildInputs_MaterialWorm.addStringValueInput(commandId + '_filterWorm', 'Фильтр', '')


            groupCmdInput_MaterialWheel = tab2ChildInputs.addGroupCommandInput(commandId + '_groupMaterial',
                                                                          'Материал Колеса')
            groupCmdInput_MaterialWheel.isExpanded = False
            groupChildInputs_MaterialWheel = groupCmdInput_MaterialWheel.children

            materialLibInputWheel = groupChildInputs_MaterialWheel.addDropDownCommandInput(commandId + '_materialLibWheel', 'Библиотека материалов', adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            listItemsWheel = materialLibInputWheel.listItems
            materialLibNamesWheel = getMaterialLibNames()
            for materialName in materialLibNamesWheel :
                listItemsWheel.add(materialName, False, '')
            listItemsWheel[0].isSelected = True
            materialListInputWheel = groupChildInputs_MaterialWheel.addDropDownCommandInput(commandId + '_materialListWheel', 'Материал', adsk.core.DropDownStyles.TextListDropDownStyle)
            materialsWheel = getMaterialsFromLib(materialLibNamesWheel[0], '')
            listItemsWheel = materialListInputWheel.listItems
            for materialName in materialsWheel :
                listItemsWheel.add(materialName, False, '')
            listItemsWheel[0].isSelected = True
            filter_materialGear = groupChildInputs_MaterialWheel.addStringValueInput(commandId + '_filterWheel', 'Фильтр', '')


            groupCmdInput_Material = tab2ChildInputs.addGroupCommandInput(commandId + '_groupMaterial',
                                                                          'Характеристики материалов')
            groupCmdInput_Material.isExpanded = False
            groupChildInputs_Material = groupCmdInput_Material.children

            Sn = groupChildInputs_Material.addValueInput('Model', 'Предел устал. прочности изгиба (Sn)', 'MPa',
                                                    adsk.core.ValueInput.createByString('165.0'))
            Sn.tooltip = "Предел устал. прочности изгиба (Sn)"
                                                
            Kw = groupChildInputs_Material.addValueInput('Model', 'Контактная усталостная прочность (Kw)', 'MPa',
                                                    adsk.core.ValueInput.createByString('0.6'))
            Kw.tooltip = "Контактная усталостная прочность (Kw)"

            Elastic = groupChildInputs_Material.addValueInput('Model', 'Модуль упругости (E)', 'MPa',
                                                    adsk.core.ValueInput.createByString('206000.0'))
            Elastic.tooltip = "Модуль упругости (E)"

            Puasson =  groupChildInputs_Material.addValueInput('Model', 'Коэффициент Пуассона (μ)', '',
                                                    adsk.core.ValueInput.createByReal(0.3))
            Puasson.tooltip = "Коэффициент Пуассона (μ)"

            Kmat = groupChildInputs_Material.addValueInput('Model', 'Коэффициент материала червяка (Kmat)', '',
                                                    adsk.core.ValueInput.createByReal(1.0))
            Kmat.tooltip = "Коэффициент материала червяка (Kmat)"

            y_Luis = groupChildInputs_Material.addValueInput('Model', 'Коэффициент Льюиса (y)', '',
                                                    adsk.core.ValueInput.createByReal(0.125))
            y_Luis.tooltip = "Коэффициент Льюиса (y)"

            groupCmdInput_Results = tab2ChildInputs.addGroupCommandInput(commandId + '_groupResults', 'Результаты')
            groupCmdInput_Results.isExpanded = False
            groupChildInputs_Results = groupCmdInput_Results.children

            # Таблица Результата Расчета
            # General_Results
            table = adsk.core.TableCommandInput.cast(
                groupChildInputs_Results.addTableCommandInput('table0', 'Inputs', 2, '1:1'))
            table.minimumVisibleRows = 2
            table.maximumVisibleRows = 3
            table.columnSpacing = 1
            table.rowSpacing = 1
            table.tablePresentationStyle = adsk.core.TablePresentationStyles.itemBorderTablePresentationStyle
            table.hasGrid = True
            # 1
            text = tab2ChildInputs.addStringValueInput('text0', 'Fr',
                                                       'Радиальная сила (Fr)')
            text.isReadOnly = True
            table.addCommandInput(text, 0, 0, False, False)

            Frad_tab = tab2ChildInputs.addTextBoxCommandInput('text0', 'TextFr',
                                                       '', 1, True)
            Frad_tab.isReadOnly = True
            table.addCommandInput(Frad_tab, 0, 1, False, False)
            # 2
            text = tab2ChildInputs.addStringValueInput('text0', 'Fn',
                                                       'Цикл нагружения (Fn)')
            text.isReadOnly = True
            table.addCommandInput(text, 1, 0, False, False)

            Fn_tab = tab2ChildInputs.addTextBoxCommandInput('text0', 'TextFn',
                                                       '', 1, True)
            Fn_tab.isReadOnly = True
            table.addCommandInput(Fn_tab, 1, 1, False, False)
            # 3
            text = tab2ChildInputs.addStringValueInput('text0', 'Vk',
                                                       'Скорость скольжения (Vk)')
            text.isReadOnly = True
            table.addCommandInput(text, 2, 0, False, False)
            
            Vk_tab = tab2ChildInputs.addTextBoxCommandInput('text0', 'TextVk',
                                                       '', 1, True)
            Vk_tab.isReadOnly = True
            Vk_tab.text = ('%.4f' % vk) + ' м/c'
            table.addCommandInput(Vk_tab, 2, 1, False, False)

            # Results_Worm
            groupChildInputs_Results.addTextBoxCommandInput(commandId + '_textBox', '',
                                                            'Червяк', 1, True)
            table = adsk.core.TableCommandInput.cast(
                groupChildInputs_Results.addTableCommandInput('table1', 'Inputs', 2, '1:1'))
            table.minimumVisibleRows = 2
            table.maximumVisibleRows = 2
            table.columnSpacing = 1
            table.rowSpacing = 1
            table.tablePresentationStyle = adsk.core.TablePresentationStyles.itemBorderTablePresentationStyle
            table.hasGrid = True
            # 1
            text = tab2ChildInputs.addStringValueInput('text0', 'Ft',
                                                       'Окружная сила (Ft)')
            text.isReadOnly = True
            table.addCommandInput(text, 0, 0, False, False)

            Ft_worm_tab = tab2ChildInputs.addTextBoxCommandInput('Count', 'TextFt',
                                                       '', 1, True)
            Ft_worm_tab.isReadOnly = True
            Ft_worm_tab.text = ('%.4f' % ((2000*float(Moment_.value))/(temp_dw1))) + ' Н'
            #((2000*float(Moment_.value))/((temp_dw2*0.01)*temp_Peredat))) + ' Н' 
            table.addCommandInput(Ft_worm_tab, 0, 1, False, False)
            # 2
            text = tab2ChildInputs.addStringValueInput('text0', 'Fa',
                                                       'Осевая сила (Fa)')
            text.isReadOnly = True
            table.addCommandInput(text, 1, 0, False, False)

            Fa_worm_tab = tab2ChildInputs.addTextBoxCommandInput('Count', 'TextFa',
                                                       '', 1, True)
            Fa_worm_tab.isReadOnly = True
            table.addCommandInput(Fa_worm_tab, 1, 1, False, False)

            # Results_WormGear
            groupChildInputs_Results.addTextBoxCommandInput(commandId + '_textBox', '',
                                                            'Червячная передача', 1, True)
            table = adsk.core.TableCommandInput.cast(
                groupChildInputs_Results.addTableCommandInput('table2', 'Inputs', 2, '1:1'))
            table.minimumVisibleRows = 3
            table.maximumVisibleRows = 5
            table.columnSpacing = 1
            table.rowSpacing = 1
            table.tablePresentationStyle = adsk.core.TablePresentationStyles.itemBorderTablePresentationStyle
            table.hasGrid = True

            # 1
            text = tab2ChildInputs.addStringValueInput('text0', 'Ft',
                                                       'Окружная сила (Ft)')
            text.isReadOnly = True
            table.addCommandInput(text, 0, 0, False, False)

            Ft_WG_tab = tab2ChildInputs.addTextBoxCommandInput('Count', 'TextFt',
                                                       '', 1 , True)
            Ft_WG_tab.isReadOnly = True
            tempS = float((str(Ft_worm_tab.text).split(' '))[0])
            Ft_WG_tab.text =('%.4f' % ((tempS*temp_Peredat*float(KPD.value))*((temp_dw1)/(temp_dw2)))) + ' Н'
            Fa_worm_tab.text = Ft_WG_tab.text
            tempS = float((str(Ft_WG_tab.text).split(' '))[0])
            Fn_tab.text = ('%.4f' % ((tempS)/(math.cos(temp_Angle_prof * math.pi/180)*math.cos(float(Angle_teeth_.value))))) + ' Н'
            table.addCommandInput(Ft_WG_tab, 0, 1, False, False)
            # 2
            text = tab2ChildInputs.addStringValueInput('text0', 'Fa',
                                                       'Осевая сила (Fa)')
            text.isReadOnly = True
            table.addCommandInput(text, 1, 0, False, False)

            Fa_WG_tab = tab2ChildInputs.addTextBoxCommandInput('Count', 'TextFa',
                                                       '', 1, True)
            Fa_WG_tab.isReadOnly = True
            Fa_WG_tab.text =  Ft_worm_tab.text
            table.addCommandInput(Fa_WG_tab, 1, 1, False, False)
            # 3
            text = tab2ChildInputs.addStringValueInput('text0', 'Fd',
                                                       'Динамическая нагрузка (Fd)')
            text.isReadOnly = True
            table.addCommandInput(text, 2, 0, False, False)
           
            Fd_WG_tab = tab2ChildInputs.addTextBoxCommandInput('Count', 'TextFd',
                                                       '', 1, True )
            Fd_WG_tab.isReadOnly = True
            tempS = float((str(Ft_WG_tab.text).split(' '))[0])
            Kv = (1200+v2)/(1200)
            Fd_WG_tab.text = ('%.4f' % (tempS*Kv)) + ' Н'
            table.addCommandInput(Fd_WG_tab, 2, 1, False, False)
            # 4
            text = tab2ChildInputs.addStringValueInput('text0', 'Fw',
                                                       'Поверхн. устал. пред. нагрузки (Fw)')
            text.isReadOnly = True
            table.addCommandInput(text, 3, 0, False, False)
            
            Fw_WG_tab = tab2ChildInputs.addTextBoxCommandInput('Count', 'TextFw',
                                                       '', 1 , True)
            Fw_WG_tab.isReadOnly = True 
            Fw_WG_tab.text = ('%.4f' % (temp_d2*0.001 * float(Width_WG.value)*0.01 * float(Kw.value)*100)) + ' Н'
            table.addCommandInput(Fw_WG_tab, 3, 1, False, False)
            # 5
            text = tab2ChildInputs.addStringValueInput('text0', 'Fs',
                                                       'Усталость изгиба пред. нагрузки (Fs)')
            text.isReadOnly = True
            table.addCommandInput(text, 4, 0, False, False)
            
            Fs_WG_tab = tab2ChildInputs.addTextBoxCommandInput('Count', 'TextFs',
                                                       '', 1, True)
            Fs_WG_tab.isReadOnly = True
            Fs_WG_tab.text = ('%.4f' % ((float(Sn.value)/10)*float(Width_WG.value)*0.01*(math.pi*temp_Module)*float(y_Luis.value))) + ' Н'
            table.addCommandInput(Fs_WG_tab, 4, 1, False, False)

            temp_Ft_worm = float(str(Ft_worm_tab.text).split(' ')[0])
            Frad_tab.text = ('%.4f' % (temp_Ft_worm * (math.tan(temp_Angle_prof * math.pi/180)*math.cos(phiz * math.pi/180))/(math.sin((tgy+phiz)*math.pi/180)))) + ' Н'        
          

            # Connect to the command related events.
            onExecute = GearCommandExecuteHandler()
            cmd.execute.add(onExecute)
            _handlers.append(onExecute)        
            
            onInputChanged = GearCommandInputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            _handlers.append(onInputChanged)     

            _app.unregisterCustomEvent('TestTrucCustomEvent')
            customEvent = _app.registerCustomEvent('TestTrucCustomEvent')
            onCustomEvent = WormHandler()
            onCustomEvent.cmd = cmd
            customEvent.add(onCustomEvent)
            _handlers.append(onCustomEvent)

        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# Event handler for the execute event.
class GearCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            global ModuleExp, NumVitkovExp, PressureAngleExp, DiamWormExp, LengthWormExp, materialWorm
            eventArgs = adsk.core.CommandEventArgs.cast(args)
            command = args.firingEvent.sender
            inputs = command.commandInputs

            des = adsk.fusion.Design.cast(_app.activeProduct)
            attribs = des.attributes
            attribs.add('WormGear', 'peredatNumber', Peredat_.selectedItem.name)
            attribs.add('WormGear', 'module', Module_.selectedItem.name)
            attribs.add('WormGear', 'pressureAngle', Angle_prof_.selectedItem.name)
            attribs.add('Worm', 'BuildWorm', selectCreateWorm.selectedItem.name)
            attribs.add('Worm', 'NumVitkov', str(Num_of_vit_worm.value))
            attribs.add('Worm', 'KolOborotov', str(KolOborotov_worm.value))
            attribs.add('Worm', 'AverageDiam', str(Av_diam_worm.value))
            attribs.add('Gear', 'BuildGear', selectCreateGear.selectedItem.name)
            attribs.add('Gear', 'numTeeth', str(Teeth_Num_Gear.value))
            attribs.add('Gear', 'thickness', str(Width_WG.value))
            attribs.add('Gear', 'koefSmesh', str(Koef_smesh_Gear.value))
            attribs.add('Gear', 'NaprVitkov', buttonRowInput.selectedItem.name)
            attribs.add('Gear', 'holediam', str(hole_diameter.value))


            if (selectCreateGear.selectedItem.name == 'Построить 3D модель'):
                des = adsk.fusion.Design.cast(_app.activeProduct)

                profile_angle = float((Angle_prof_.selectedItem.name).split(' ')[0]) * math.pi/180 
                diamAv = float((_D_WG_tab.text).split(' ')[0])
                Module_creation = float((Module_.selectedItem.name).split(' ')[0])
                diaPitch = 25.4 / Module_creation
                Teeth_NumGear = float(Teeth_Num_Gear.value)
                widthGear = Width_WG.value
                helixangle = float(Angle_teeth_.value)
                holeDiam = float(hole_diameter.value)

                for input in inputs:
                    if input.id == commandId + '_materialListWheel':
                        materialListInputWheel = input
                materialWheel = getMaterial(materialListInputWheel.selectedItem.name)

                gearComp = drawGear(des, diaPitch, Teeth_NumGear, widthGear, profile_angle, Module_creation, helixangle, holeDiam, materialWheel)
            
            if (selectCreateWorm.selectedItem.name == 'Построить 3D модель'):
                ModuleExp = str((_Module_tab.text).split(' ')[0]) + ' mm'
                NumVitkovExp = str(Num_of_vit_worm.value)
                PressureAngleExp = str(Angle_prof_.selectedItem.name)
                DiamWormExp = str(('%.4f' %(Av_diam_worm.value)) + ' mm')
                LengthWormExp = str(int(KolOborotov_worm.value))
                
                for input in inputs:
                    if input.id == commandId + '_materialListWorm':
                        materialListInputWorm = input
                materialWorm = getMaterial(materialListInputWorm.selectedItem.name)
                
                _app.fireCustomEvent('TestTrucCustomEvent')
            


        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
        

class WormHandler(adsk.core.CustomEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            importManager = _app.importManager
            rootComp = _app.activeProduct.rootComponent



            filename = str(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'resources\Models\Worm.f3d'))

            importOptions = importManager.createFusionArchiveImportOptions(filename)
            importManager.importToTarget(importOptions, rootComp)
            

            pulleyOccurance = rootComp.occurrences.item(rootComp.occurrences.count-1)

            pulleyOccurance.component.material = materialWorm
            parametersWorm = pulleyOccurance.component.parentDesign.allParameters
            ModuleExpression = parametersWorm.itemByName('Module')
            NumVitkov = parametersWorm.itemByName('NumberOfStarts')
            PressureAngle = parametersWorm.itemByName('PressureAngle')
            DiamWorm = parametersWorm.itemByName('Diameter')
            LengthWorm = parametersWorm.itemByName('Length')
            ModuleExpression.expression = ModuleExp
            NumVitkov.expression = NumVitkovExp
            PressureAngle.expression = PressureAngleExp
            DiamWorm.expression = DiamWormExp
            LengthWorm.expression = LengthWormExp

            
        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
        

def drawGear(design, diametralPitch, numTeeth, thickness, pressureAngle, module, helixAngle, hole_diam, materialWheel):
    try:
        # The diametral pitch is specified in inches but everthing
        # here expects all distances to be in centimeters, so convert
        # for the gear creation.
        diametralPitch = diametralPitch /2.54
        thickness = thickness * 0.1
        hole_diam = hole_diam * 0.1

        pitchDia = numTeeth / diametralPitch

        #addendum = 1.0 / diametralPitch
        if (diametralPitch < (20 *(math.pi/180))-0.000001):
            dedendum = 1.157 / diametralPitch
        else:
            circularPitch = math.pi / diametralPitch
            if circularPitch >= 20:
                dedendum = 1.25 / diametralPitch
            else:
                dedendum = (1.2 / diametralPitch) + (.002 * 2.54)                

        rootDia = pitchDia - (2 * dedendum)
        
        baseCircleDia = pitchDia * math.cos(pressureAngle)
        outsideDia = (numTeeth + 2) / diametralPitch
        
        # Create a new component by creating an occurrence.
        occs = design.rootComponent.occurrences
        mat = adsk.core.Matrix3D.create()
        newOcc = occs.addNewComponent(mat)        
        newComp = adsk.fusion.Component.cast(newOcc.component)
        
        # Create a new sketch.
        sketches = newComp.sketches
        newComp.name
        xyPlane = newComp.xYConstructionPlane
        baseSketch = sketches.add(xyPlane)
        positionX = -(((float(Av_diam_worm.value) / 2)) + ((numTeeth * module)/2))*0.1
        positionY = ((int(KolOborotov_worm.value) * 0.5)*(module*math.pi))*0.1
        # Draw a circle for the base.
        baseSketch.sketchCurves.sketchCircles.addByCenterRadius(adsk.core.Point3D.create(0, 0, 0), rootDia/2.0)
        
        # Create a second sketch for the tooth.
        toothSketch = sketches.add(xyPlane)
        for number in range (int(numTeeth)):
            # Calculate points along the involute curve.
            originPoint = adsk.core.Point3D.create(0, 0, 0)
            involutePointCount = 15 
            involuteIntersectionRadius = baseCircleDia / 2.0
            involutePoints = []
            keyPoints = []
            involuteSize = (outsideDia - baseCircleDia) / 2.0
            for i in range(0, involutePointCount):
                involuteIntersectionRadius = (baseCircleDia / 2.0) + ((involuteSize / (involutePointCount - 1)) * i)
                newPoint = involutePoint(baseCircleDia / 2.0, involuteIntersectionRadius)
                involutePoints.append(newPoint)
                
            # Get the point along the tooth that's at the pictch diameter and then
            # calculate the angle to that point.
            pitchInvolutePoint = involutePoint(baseCircleDia / 2.0, pitchDia / 2.0)
            pitchPointAngle = math.atan(pitchInvolutePoint.y / pitchInvolutePoint.x)

            # Determine the angle defined by the tooth thickness as measured at
            # the pitch diameter circle.
            toothThicknessAngle = (2 * math.pi) / (2 * numTeeth)
            
            # Determine the angle needed for the specified backlash.
            backlashAngle = (0 / (pitchDia / 2.0)) * .25
            
            # Determine the angle to rotate the curve.
            rotateAngle = -((toothThicknessAngle/2) + pitchPointAngle - backlashAngle)
            
            # Rotate the involute so the middle of the tooth lies on the x axis.
            cosAngle = math.cos(rotateAngle)
            sinAngle = math.sin(rotateAngle)
            for i in range(0, involutePointCount):
                x = involutePoints[i].x
                y = involutePoints[i].y
                involutePoints[i].x = x * cosAngle - y * sinAngle
                involutePoints[i].y = x * sinAngle + y * cosAngle

            # Create a new set of points with a negated y.  This effectively mirrors the original
            # points about the X axis.
            involute2Points = []
            for i in range(0, involutePointCount):
                involute2Points.append(adsk.core.Point3D.create(involutePoints[i].x, -involutePoints[i].y, 0))

            # Rotate involute
            rotation = (number / int(numTeeth)) * 2 * math.pi
            if rotation:
                cosAngle = math.cos(rotation)
                sinAngle = math.sin(rotation)
                for i in range(0, involutePointCount):
                    x = involutePoints[i].x
                    y = involutePoints[i].y
                    involutePoints[i].x = x * cosAngle - y * sinAngle
                    involutePoints[i].y = x * sinAngle + y * cosAngle
                    x = involute2Points[i].x
                    y = involute2Points[i].y
                    involute2Points[i].x = x * cosAngle - y * sinAngle
                    involute2Points[i].y = x * sinAngle + y * cosAngle

            curve1Angle = math.atan2(involutePoints[0].y, involutePoints[0].x)
            curve2Angle = math.atan2(involute2Points[0].y, involute2Points[0].x)
            if curve2Angle < curve1Angle:
                curve2Angle += math.pi * 2
            toothSketch.isComputeDeferred = True
            
            # Create and load an object collection with the points.
            pointSet1 = adsk.core.ObjectCollection.create()
            pointSet2 = adsk.core.ObjectCollection.create()
            for i in range(0, involutePointCount):
                pointSet1.add(involutePoints[i])
                pointSet2.add(involute2Points[i])

            midIndex = int(pointSet1.count / 2)
            keyPoints.append(pointSet1.item(0))
            keyPoints.append(pointSet2.item(0))
            keyPoints.append(pointSet1.item(midIndex))
            keyPoints.append(pointSet2.item(midIndex))

            # Create splines.
            spline1 = toothSketch.sketchCurves.sketchFittedSplines.add(pointSet1)
            spline2 = toothSketch.sketchCurves.sketchFittedSplines.add(pointSet2)
            oc = adsk.core.ObjectCollection.create()
            oc.add(spline2)
            (_, _, crossPoints) = spline1.intersections(oc)
            assert len(crossPoints) == 0 or len(crossPoints) == 1, 'Failed to compute a valid involute profile!'
            if len(crossPoints) == 1:
                # involute splines cross, clip the tooth
                # clip = spline1.endSketchPoint.geometry.copy()
                # spline1 = spline1.trim(spline2.endSketchPoint.geometry).item(0)
                # spline2 = spline2.trim(clip).item(0)
                keyPoints.append(crossPoints[0])
            else:
                # Draw the tip of the tooth - connect the splines
                if numTeeth >= 100:
                    toothSketch.sketchCurves.sketchLines.addByTwoPoints(spline1.endSketchPoint, spline2.endSketchPoint)
                    keyPoints.append(spline1.endSketchPoint.geometry)
                    keyPoints.append(spline2.endSketchPoint.geometry)
                else:
                    tipCurve1Angle = math.atan2(involutePoints[-1].y, involutePoints[-1].x)
                    tipCurve2Angle = math.atan2(involute2Points[-1].y, involute2Points[-1].x)
                    if tipCurve2Angle < tipCurve1Angle:
                        tipCurve2Angle += math.pi * 2
                    tipRad = originPoint.distanceTo(involutePoints[-1])
                    tipArc = toothSketch.sketchCurves.sketchArcs.addByCenterStartSweep(
                        originPoint,
                        adsk.core.Point3D.create(math.cos(tipCurve1Angle) * tipRad,
                                                math.sin(tipCurve1Angle) * tipRad,
                                                0),
                        tipCurve2Angle - tipCurve1Angle)
                    keyPoints.append(tipArc.startSketchPoint.geometry)
                    keyPoints.append(adsk.core.Point3D.create(tipRad, 0, 0))
                    keyPoints.append(tipArc.endSketchPoint.geometry)

            # Draw root circle
            # rootCircle = sketch.sketchCurves.sketchCircles.addByCenterRadius(originPoint, self.gear.rootDiameter/2)
            rootArc = toothSketch.sketchCurves.sketchArcs.addByCenterStartSweep(
                originPoint,
                adsk.core.Point3D.create(math.cos(curve1Angle) * (rootDia / 2 - 0.01),
                                        math.sin(curve1Angle) * (rootDia / 2 - 0.01),
                                        0),
                curve2Angle - curve1Angle)

            # if the offset tooth profile crosses the offset circle then trim it, else connect the offset tooth to the circle
            oc = adsk.core.ObjectCollection.create()
            oc.add(spline1)
            if True:
                if rootArc.intersections(oc)[1].count > 0:
                    spline1 = spline1.trim(originPoint).item(0)
                    spline2 = spline2.trim(originPoint).item(0)
                    rootArc.trim(rootArc.startSketchPoint.geometry)
                    rootArc.trim(rootArc.endSketchPoint.geometry)
                else:
                    toothSketch.sketchCurves.sketchLines.addByTwoPoints(originPoint, spline1.startSketchPoint).trim(
                        originPoint)
                    toothSketch.sketchCurves.sketchLines.addByTwoPoints(originPoint, spline2.startSketchPoint).trim(
                        originPoint)
            else:
                if rootArc.intersections(oc)[1].count > 0:
                    spline1 = spline1.trim(originPoint).item(0)
                    spline2 = spline2.trim(originPoint).item(0)
                rootArc.deleteMe()
                sketch.sketchCurves.sketchLines.addByTwoPoints(originPoint, spline1.startSketchPoint)
                sketch.sketchCurves.sketchLines.addByTwoPoints(originPoint, spline2.startSketchPoint)

        ### Extrude the tooth.
        # Get the profile defined by the tooth.
        toothSketch.isComputeDeferred = False
        
        profs = adsk.core.ObjectCollection.create()
        # Add all of the profiles to the collection.
        for prof in baseSketch.profiles:
            profs.add(prof)
        for prof in toothSketch.profiles:
            profs.add(prof)
        # Create a new sketch.
        sketches = newComp.sketches
        xzPlane = newComp.xZConstructionPlane
        sketchVertical = sketches.add(xzPlane)

        # Draw a circle for the base.
        point1 = adsk.core.Point3D.create(0,0,0)
        point2 = adsk.core.Point3D.create(0,thickness,0)
        linesweep = sketchVertical.sketchCurves.sketchLines.addByTwoPoints(point1, point2)
       
        path = newComp.features.createPath(linesweep)

        sweeps = newComp.features.sweepFeatures
        sweepInput = sweeps.createInput(profs, path, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
        if buttonRowInput.selectedItem.name == 'Левое':
            twist = -(thickness / (math.tan(math.radians(90) + helixAngle*math.pi/180) * (pitchDia/2)))
        elif buttonRowInput.selectedItem.name == 'Правое':
            twist = (thickness / (math.tan(math.radians(90) + helixAngle*math.pi/180) * (pitchDia/2)))
        sweepInput.twistAngle = adsk.core.ValueInput.createByReal(twist)
        sweep = sweeps.add(sweepInput)
    
        
        # Add an attribute to the component with all of the input values.  This might 
        # be used in the future to be able to edit the gear.     
        gearValues = {}
        gearValues['diametralPitch'] = str(diametralPitch * 2.54)
        gearValues['numTeeth'] = str(numTeeth)
        gearValues['thickness'] = str(thickness)
        gearValues['pressureAngle'] = str(pressureAngle)
        attrib = newComp.attributes.add('Wheel', 'Values',str(gearValues))
        newComp.name = 'Wheel'
        
        count = occs.count
        subComp1 = occs.item(count-1).component
       
        # Rotate the wheel
        bodyComp = subComp1
        body = bodyComp.bRepBodies.item(0)
        inputEnts = adsk.core.ObjectCollection.create()
        inputEnts.add(body)


        trans = adsk.core.Matrix3D.create()

        # transform to worm
        vector = adsk.core.Vector3D.create(positionX, positionY, thickness/2)
        trans.translation = vector    

        # rotations
        roty = adsk.core.Matrix3D.create()
        angle = 180
        roty.setToRotation(math.radians(angle), newComp.yConstructionAxis.geometry.getData()[2], adsk.core.Point3D.create(0, 0, 0))
        trans.transformBy(roty)
            
        rotz = adsk.core.Matrix3D.create()
        angle =  ( 360 / ( numTeeth * 2 ) ) + ( (twist*180/math.pi) / 2 ) 
        rotz.setToRotation(-math.radians(angle), newComp.zConstructionAxis.geometry.getData()[2], adsk.core.Point3D.create(-positionX,positionY,-thickness))
        trans.transformBy(rotz)
            
            
        ents = adsk.core.ObjectCollection.create()
        ents.add(body)
        moveInput = bodyComp.features.moveFeatures.createInput(ents, trans)
        bodyComp.features.moveFeatures.add(moveInput)

        applyMaterialToEntities(materialWheel, ents)
        
        if hole_diam > 0:
            holeSketch = sketches.add(xyPlane)
            profile = adsk.fusion.Profile.cast(None)
            holeSketch.sketchCurves.sketchCircles.addByCenterRadius(adsk.core.Point3D.create(-positionX, positionY, thickness/2), hole_diam/2.0)
            profile = holeSketch.profiles.item(0)
            
            #### Extrude the circle to create the base of the gear.
            extrudes = newComp.features.extrudeFeatures
            extInput = extrudes.createInput(profile, adsk.fusion.FeatureOperations.CutFeatureOperation)

            distance = adsk.core.ValueInput.createByReal(-thickness)
            extInput.setDistanceExtent(False, distance)

            # Create the extrusion.
            baseExtrude = extrudes.add(extInput)

        return newComp
    except Exception as error:
        _ui.messageBox("draw Wheel Failed : " + str(error)) 
        return None

# Calculate points along an involute curve.
def involutePoint(baseCircleRadius, distFromCenterToInvolutePoint):
    try:  
        l = math.sqrt(
            distFromCenterToInvolutePoint * distFromCenterToInvolutePoint - baseCircleRadius * baseCircleRadius)
        alpha = l / baseCircleRadius
        theta = alpha - math.acos(baseCircleRadius / distFromCenterToInvolutePoint)
        x = distFromCenterToInvolutePoint * math.cos(theta)
        y = distFromCenterToInvolutePoint * math.sin(theta)
        return adsk.core.Point3D.create(x, y, 0)
    except:
        if _ui:
            _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# Event handler for the inputChanged event.
class GearCommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            eventArgs = adsk.core.InputChangedEventArgs.cast(args)
            changedInput = eventArgs.input
            
            global _units, Peredat, commandId
            
            cmd = args.firingEvent.sender
            inputs = cmd.commandInputs
            materialListInput = None
            filterInput = None
            materialLibInput = None
            for inputI in inputs:
                if inputI.id == commandId + '_materialListWorm':
                    materialListInput = inputI
                elif inputI.id == commandId + '_filterWorm':
                    filterInput = inputI
                elif inputI.id == commandId + '_materialLibWorm':
                    materialLibInput = inputI
            cmdInput = args.input
            if cmdInput.id == commandId + '_materialLibWorm' or cmdInput.id == commandId + '_filterWorm':
                materials = getMaterialsFromLib(materialLibInput.selectedItem.name, filterInput.value)
                replaceItems(materialListInput, materials)
            
            for inputI in inputs:
                if inputI.id == commandId + '_materialListWheel':
                    materialListInput = inputI
                elif inputI.id == commandId + '_filterWheel':
                    filterInput = inputI
                elif inputI.id == commandId + '_materialLibWheel':
                    materialLibInput = inputI
            cmdInput = args.input
            if cmdInput.id == commandId + '_materialLibWheel' or cmdInput.id == commandId + '_filterWheel':
                materials = getMaterialsFromLib(materialLibInput.selectedItem.name, filterInput.value)
                replaceItems(materialListInput, materials)




            if buttonSaveLoad.listItems[0].isSelected == True:
                root = tk.Tk()
                root.withdraw()
                folder_path = filedialog.asksaveasfile(initialfile = 'data.json',
                                            initialdir= str(os.path.dirname(os.path.realpath(__file__))),
                                            title="Сохранить данные",
                                            filetypes=(("JSON files", '*.json'),("all files", "*.*")),
                                            mode='w')
                buttonSaveLoad.listItems[0].isSelected = False    
                if folder_path is None:
                    return

                params = exportParameters()
                with open(folder_path.name, 'w') as f:
                    f.write(json.dumps(params))
                _ui.messageBox('Данные успешно сохранены!')
               
            if buttonSaveLoad.listItems[1].isSelected == True:
                root = tk.Tk()
                root.withdraw()
                file_path = filedialog.askopenfilename(initialdir= str(os.path.dirname(os.path.realpath(__file__))),
                                                        title="Открыть файл",
                                                        filetypes=(("JSON files", '*.json'),("all files", "*.*")))
                buttonSaveLoad.listItems[1].isSelected = False    
                if not file_path:
                    return

                with open(file_path, 'r') as f:
                    try:
                        data = json.loads(f.read())
                    except:
                        _ui.messageBox('Ошибка чтения файла!')
                        return

                importParameters(data) 
                f.close()

             #Values from dropdowns
            Module = (Module_.selectedItem.name).split(' ')
            Module = Module[0]
            Angle_prof = (Angle_prof_.selectedItem.name).split(' ')
            Angle_prof = float(Angle_prof[0])
            Peredat = (Peredat_.selectedItem.name).split(' ')
            Peredat = float(Peredat[0])
            Length_wormNarez.isEnabled = False
            Length_wormNarez.value = int(KolOborotov_worm.value)*float(Module)*3.14

            if radio_WormSize.listItems[0].isSelected == True:
                Koef_Diam_worm.isEnabled = True
                Angle_teeth_.isEnabled = False
                Av_diam_worm.isEnabled = False
                Av_diam_worm.value = Koef_Diam_worm.value * float(Module)
                Angle_teeth_.value = math.atan(float(Num_of_vit_worm.value)/(float(Koef_Diam_worm.value)))*180/math.pi
            elif radio_WormSize.listItems[1].isSelected:
                Koef_Diam_worm.isEnabled = False
                Angle_teeth_.isEnabled = True
                Av_diam_worm.isEnabled = False
                Koef_Diam_worm.value = (float(Num_of_vit_worm.value))/(math.tan((Angle_teeth_.value*math.pi)/180))
                Av_diam_worm.value = float(Module) * float(Koef_Diam_worm.value)
            elif radio_WormSize.listItems[2].isSelected:
                Koef_Diam_worm.isEnabled = False
                Angle_teeth_.isEnabled = False
                Av_diam_worm.isEnabled = True
                Koef_Diam_worm.value = float(Av_diam_worm.value)/float(Module)
                Angle_teeth_.value = math.atan(float(Num_of_vit_worm.value)/(float(Koef_Diam_worm.value)))*180/math.pi

            if radio_CountType.listItems[0].isSelected == True:
                Teeth_Num_Gear.isEnabled = False
                Teeth_Num_Gear.value = int(Peredat * float(Num_of_vit_worm.value))  
                Peredat_.isEnabled = True
                Peredat_.isVisible = True
                Peredat_Changed.isVisible = False
            elif radio_CountType.listItems[1].isSelected == True :
                Teeth_Num_Gear.isEnabled = True
                Peredat_.isEnabled = False
                Peredat_Changed.value =  float(Teeth_Num_Gear.value / Num_of_vit_worm.value)
                Peredat_Changed.isEnabled = False
                Peredat_Changed.isVisible = True
                Peredat_.isVisible = False
                Peredat = Peredat_Changed.value

            if(hole_Check.value == False):
                hole_diameter.isVisible = False
                hole_diameter.value = 0
            else:
                hole_diameter.isVisible = True

            #Values of diameters    
            pitchDia = int(Teeth_Num_Gear.value) / ( (25.4 / float(Module)) / 2.54)
            twistAngle = (float(Width_WG.value) / (math.tan(math.radians(90) + float(Angle_teeth_.value) * math.pi/180) * (pitchDia/2)))
            d1 = float(Koef_Diam_worm.value) * float(Module)
            d2 = int(Teeth_Num_Gear.value) * float(Module)
            dw1 = d1 + 2 * float(Koef_smesh_Gear.value) * float(Module)
            dw2 = d2
            da1 = d1 + (2*float(Module)*1)
            db1 = d1 * math.cos(Angle_prof * math.pi/180)
            da2 = d2 + 2 * float(Module)*(1 + float(Koef_smesh_Gear.value))
            db2 = d2 * math.cos(Angle_prof * math.pi/180)
            #Epsilon
            Epsilon = (float(Width_WG.value)*math.sin(float(Angle_teeth_.value)))/(math.pi*float(Module))

            #Value of angle prof
            tan_on_cos = math.tan(Angle_prof * math.pi/180) * math.cos(float(Angle_teeth_.value*math.pi/180))

            #Min koef
            Ha0 = 1 + 0.3 - 0.3*(1- math.sin(Angle_prof * math.pi/180)) 
            Zv2 = float(Teeth_Num_Gear.value)/(math.cos(Angle_teeth_.value)**3)

            if(Kpd_Check.value == False):
                KPD.isVisible = False
            else:
                KPD.isVisible = True

            v1 = (math.pi*d1*float(Velocity_.value))/(60000)
            tgy = math.atan(float(Num_of_vit_worm.value)/float(Koef_Diam_worm.value))*180/math.pi
            v2 = v1*(float(Num_of_vit_worm.value)/float(Koef_Diam_worm.value)) 
            vk = v1/math.cos(tgy*math.pi/180)
            phiz = math.atan(0.02+(0.03)/(vk))*180/math.pi
            Kv = (1200+v2)/(1200)

            #Radio inputs
            #Peredacha
            Moment_.value = (60000*float(Power_.value))/(2*math.pi*float(Velocity_.value))
            Moment_WG.value = (60000*float(Power_WG.value))/(2*math.pi*float(Velocity_WG.value))
            Ft_WG = (2000*float(Moment_WG.value))/(dw2)
            Ft_worm = (2000*float(Moment_.value))/(dw1)
            radioButtonS.isVisible = False 
            if(radioButtonS.listItems[0].isSelected ==  True):
                Power_.isVisible = False
                Velocity_.isVisible = False
                Moment_.isVisible = False
                Power_WG.isVisible = True
                Velocity_WG.isVisible = True
                Moment_WG.isVisible = True
               
                Velocity_.value = float(Velocity_WG.value)*Peredat

                KPD.value = ((math.tan((tgy-phiz)*math.pi/180))/(math.tan(tgy*math.pi/180)))*0.96
                Power_.value = float(Power_WG.value) * float(KPD.value)
                
                Moment_.value = (float(Moment_WG.value)/(Peredat))*float(KPD.value)
                Ft_worm = (2000*float(Moment_.value))/(dw1)

            #Chervyak
            else:
                Power_.isVisible = True
                Velocity_.isVisible = True
                Moment_.isVisible = True
                Power_WG.isVisible = False
                Velocity_WG.isVisible = False
                Moment_WG.isVisible = False
                
                Velocity_WG.value = float(Velocity_.value)/Peredat

                KPD.value = ((math.tan(tgy*math.pi/180))/(math.tan((tgy+phiz)*math.pi/180)))*0.96
                Power_WG.value = float(Power_.value) * float(KPD.value)
                
                Moment_WG.value = float(KPD.value)*float(Moment_.value)*Peredat
                Ft_WG = (2000*float(Moment_WG.value))/(dw2)
                
            if changedInput.id == 'Model':
                #Table Model 
                _Aw_tab.text = ('%.4f' %((dw1+dw2)/2)) + ' мм'
                Eps_tab.text = ('%.4f' % Epsilon)
                _Module_tab.text = Module + ' мм'
                _Alpha_tab.text = ('%.4f' % (math.atan(tan_on_cos) *180/math.pi)) + ' deg'

                _naruz_Diam_tab.text = ('%.4f' %(da1)) + ' мм'    
                _Dsr_tab.text = ('%.4f' %(Av_diam_worm.value)) + ' мм'
                _Df_tab.text = ('%.4f' %(d1-2*float(Module)*(1+0.3))) + ' мм'

                _Da_WG_tab.text = str(d2+2*float(Module)*(1 + float(Koef_smesh_Gear.value))) + ' мм'
                _D_WG_tab.text = str(int(Teeth_Num_Gear.value)*float(Module)) + ' мм'
                _Df_WG_tab.text = str(d2 - 2 * float(Module)*(1 + 0.3 - float(Koef_smesh_Gear.value))) + ' мм' 
                # _xmin_tab.text = ('%.4f' % (Ha0-((Zv2/2)*(1+(0.2723/1))*(math.sin(Angle_prof * math.pi/180))**2)))

                #Table Count
                Frad_tab.text = ('%.4f' % (Ft_worm*(math.tan(Angle_prof * math.pi/180)*math.cos(phiz * math.pi/180))/(math.sin((tgy+phiz)*math.pi/180)))) + ' Н'
                Fn_tab.text = ('%.4f' % ((Ft_WG)/(math.cos(Angle_prof * math.pi/180)*math.cos(float(Angle_teeth_.value))))) + ' Н'
                Vk_tab.text = ('%.4f' % (vk)) + ' м/c'
                
                Ft_worm_tab.text = ('%.4f' % (Ft_worm)) + ' Н'
                Fa_worm_tab.text =('%.4f' % (Ft_WG)) + ' Н'
                Fa_WG_tab.text = Ft_worm_tab.text
                Ft_WG_tab.text  = Fa_worm_tab.text
            
                Ft_WG_val = float((str(Fa_worm_tab.text).split(' '))[0])
                Fd_WG_tab.text = ('%.4f' % (Ft_WG_val*Kv)) + ' Н'
                Fw_WG_tab.text = ('%.4f' % (d2*0.001 * float(Width_WG.value)*0.01 * float(Kw.value)*100)) + ' Н'
                Fs_WG_tab.text = ('%.4f' % ((float(Sn.value)/10)*float(Width_WG.value)*0.01*(math.pi*float(Module))*float(y_Luis.value))) + ' Н'
                
            if buttonimportParams.listItems[0].isSelected == True:
                generatePdfTable()
                buttonimportParams.listItems[0].isSelected = False
               
            if buttonimportParams.listItems[1].isSelected == True:
                generateWordTable()
                buttonimportParams.listItems[1].isSelected = False
                

                #Epsilons
                # pb = math.pi * math.cos(Angle_prof * math.pi/180)
                # pn = math.pi*(float(Module)/math.cos( float(Angle_teeth_.value)))
                # Alpha_t  = math.atan((math.tan(20*math.pi/180))/(math.cos(float(Angle_teeth_.value)))) * 180/math.pi
                # Beta = math.acos((float(Module)*(float(Teeth_Num_Gear.value+float(Num_of_vit_worm.value))))/(2*((dw1+dw2)/2)))*180/math.pi
                # EA = ( float(Num_of_vit_worm.value) * math.tan(Alpha_a1 * math.pi/180) + float(Teeth_Num_Gear.value) 
                #                                 * math.tan(Alpha_a2*math.pi/180) -(float(Num_of_vit_worm.value) + float(Teeth_Num_Gear.value)) *math.tan(Alpha_t*math.pi/180) )/(2*math.pi)
                # EB = (float(Width_WG.value)*math.sin(float(Angle_teeth_.value)))/(math.pi*3.831)
                # Eps_a = (math.sqrt((da2**2)-(db2**2)) + (((2*float(Module))/(math.sin(Angle_prof * math.pi/180)))) - d2*math.sin(Angle_prof * math.pi/180))/(2*pb)
                # Eps_b = (float(Width_WG.value)*math.sin(float(Angle_teeth_.value)))/pn
                #Alpha_a1 = math.acos(db1/da1)*180/math.pi
                #Alpha_a2 = math.acos(db2/da2)*180/math.pi
        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def generateWordTable():

    doc = docx.Document()
    
    # Table data in a form of list
    data = generateData(isForTable=True)
    
    # Creating a table object
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'

    row = table.rows[0].cells
    row[0].text = 'Параметр'
    row[1].text = 'Значение'
    # Adding data from the list to the table
    for id, name in data:
        # Adding a row and then adding data in it.
        row = table.add_row().cells
        # Converting id to string as table can only take string input
        row[0].text = str(id)
        row[1].text = name
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(initialfile = 'WormGearParameters.docx',
                                            initialdir= str(os.path.dirname(os.path.realpath(__file__))),
                                            title="Сохранить файл",
                                            filetypes=(("Документ Word", '*.docx'),("all files", "*.*")))
    if not file_path:
        return
    if '.docx' in file_path:
        doc.save(file_path)
    else:
        doc.save(file_path + '.docx')
    
    _ui.messageBox('Файл создан!')


def generatePdfTable():

    data = generateData(isForTable=True)
    
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '', str(os.path.dirname(os.path.realpath(__file__)))+'/fonts/DejaVuSans.ttf', uni=True)
    pdf.set_font('DejaVu', '', 13)
    line_height = pdf.font_size * 2.5
    col_width = pdf.epw / 2  
    for row in data:
        for datum in row:
            pdf.multi_cell(col_width, line_height, datum, border=1, ln=3, max_line_height=pdf.font_size)
        pdf.ln(line_height)
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(initialfile = 'WormGearParameters.pdf',
                                            initialdir= str(os.path.dirname(os.path.realpath(__file__))),
                                            title="Сохранить файл",
                                            filetypes=(("PDF", '*.pdf'),("all files", "*.*")))
    if not file_path:
        return
    if '.pdf' in file_path:
        pdf.output(file_path)
    else:
        pdf.output(file_path + '.pdf')
    
    _ui.messageBox('Файл создан!')


def generateData(isForTable):
    global Peredat
    if radio_CountType.listItems[0].isSelected == True:
        Peredat = Peredat_.selectedItem.name

    if isForTable == False:
        params = {'initial parameter': str(radio_CountType.selectedItem.name) ,
                'worm_size': str(radio_WormSize.selectedItem.name),
                'gear_ratio': Peredat,
                'module': Module_.selectedItem.name,
                'profile_angle': Angle_prof_.selectedItem.name,
                'tooth_angle': Angle_teeth_.value,
                'number_of_turns': Num_of_vit_worm.value,
                'number_of_turnsVit': KolOborotov_worm.value, 
                'diameter_factor': Koef_Diam_worm.value,
                'average_diameter': Av_diam_worm.value,
                'number_of_teeth': Teeth_Num_Gear.value,
                'gear_width': Width_WG.value,
                'bias_factor': Koef_smesh_Gear.value,
                'tooth_direction': buttonRowInput.selectedItem.name,
                'hole_diameter': hole_diameter.value,
                'power': Power_.value,
                'speed': Velocity_.value,
                'torque': Moment_.value,
                'tensile_strength': Sn.value,
                'contact_strength': Kw.value,
                'elastic_modulus': Elastic.value,
                'Poisson_ratio': Puasson.value,
                'worm_material_factor': Kmat.value,
                'Lewis_ratio': y_Luis.value
        }
    else:
        params = (
                ("Передаточное отношение", str(Peredat)),
                ('Модуль', str(Module_.selectedItem.name)),
                ('Угол профиля, мм', str(Angle_prof_.selectedItem.name)),
                ('Угол наклона зуба, °', str(Angle_teeth_.value)),
                ('Межосевое расстояние, мм', str(_Aw_tab.text)),
                ('Коэффициент осевого перекрытия', str(Eps_tab.text)),
                ('Угол профиля, °',  str(_Alpha_tab.text)),
                ('Количество витков', str(Num_of_vit_worm.value)),
                ('Количество оборотов витков', str(KolOborotov_worm.value)),
                ('Длина нарезной части червяка, мм', str(Length_wormNarez.value)),
                ('Коэффициент диаметра', str(Koef_Diam_worm.value)),
                ('Наружный диаметр червяка',  str(_naruz_Diam_tab.text)),
                ('Средний диаметр червяка', str(_Dsr_tab.text)),
                ('Диаметр впадин червяка', str(_Df_tab.text)), 
                ('Количество зубьев колеса', str(Teeth_Num_Gear.value)),
                ('Ширина червячного колеса, мм', str(Width_WG.value)),
                ('Коэффициент смещения', str(Koef_smesh_Gear.value)),
                ('Направление зубьев', str(buttonRowInput.selectedItem.name)),
                ('Диаметр отверстия колеса, мм', str(hole_diameter.value)),
                ('Наружный диаметр колеса', str(_Da_WG_tab.text)),
                ('Средний диаметр колеса',str( _D_WG_tab.text)),
                ('Диаметр впадин колеса', str(_Df_WG_tab.text)), 
                ('Мощность, кВт', str(Power_.value)),
                ('Скорость, об/мин', str(Velocity_.value)),
                ('Крутящий момент, Нм', str(Moment_.value)),
                ('КПД', str(KPD.value)),
                ("Предел устал. прочности изгиба (Sn), Па", str( Sn.value)),
                ("Контактная усталостная прочность (Kw), Па", str(Kw.value)),
                ("Модуль упругости (E), Па", str(Elastic.value)),
                ("Коэффициент Пуассона (μ)", str(Puasson.value)),
                ("Коэффициент материала червяка (Kmat)", str(Kmat.value)),
                ("Коэффициент Льюиса (y)", str(y_Luis.value)), 
                ("Радиальная сила (Fr)", str(Frad_tab.text)),
                ("Цикл нагружения (Fn)", str(Fn_tab.text)),
                ("Скорость скольжения (vk)", str(Vk_tab.text)),
                ("Окружная сила червяка (Ft)", str(Ft_worm_tab.text)),
                ("Осевая сила червяка (Fa)", str(Fa_worm_tab.text)),
                ("Окружная сила колеса (Ft)", str(Fa_WG_tab.text)),
                ("Осевая сила колеса (Fa)", str(Ft_WG_tab.text)),
                ("Динамическая нагрузка (Fd)", str(Fd_WG_tab.text)),
                ("Поверхн. устал. пред. нагрузки (Fw)", str(Fw_WG_tab.text)),
                ("Усталость изгиба пред. нагрузки (Fs)", str(Fs_WG_tab.text)),
                )
    return params


def importParameters(data):
    try:
        if data['initial parameter']  == 'Передаточное отношение':
            radio_CountType.listItems[0].isSelected = True
            for element in Peredat_.listItems:
                if element.name == data['gear_ratio']:
                    element.isSelected = True
        elif data['initial parameter']  == 'Количество зубьев':
            radio_CountType.listItems[1].isSelected = True

        if data['worm_size']  == 'Коэффициент диаметра':
            radio_WormSize.listItems[0].isSelected = True
        elif data['worm_size']  == 'Угол наклона зуба':
            radio_WormSize.listItems[1].isSelected = True
        elif data['worm_size']  == 'Средний диаметр':
            radio_WormSize.listItems[2].isSelected = True
        
        for element in Module_.listItems:
            if element.name == data['module']:
                element.isSelected = True

        for element in Angle_prof_.listItems:
            if element.name == data['profile_angle']:
                element.isSelected = True
                
        for element in buttonRowInput.listItems:
            if element.name == data['tooth_direction']:
                element.isSelected = True
        
        if data['hole_diameter'] != 0:
            hole_Check.value = True
            hole_diameter.value = float(data['hole_diameter'])
        
        Angle_teeth_.value = float(data['tooth_angle'])
        Num_of_vit_worm.value = int(data['number_of_turns'])
        KolOborotov_worm.value = int(data['number_of_turnsVit'])
        Koef_Diam_worm.value = float(data['diameter_factor'])
        Av_diam_worm.value = float(data['average_diameter'])
        Teeth_Num_Gear.value = int(data['number_of_teeth'])
        Width_WG.value = int(data['gear_width']) 
        Koef_smesh_Gear.value = float(data['bias_factor'])
        Power_.value = float(data['power'])
        Velocity_.value = float(data['speed'])
        Moment_.value = float(data['torque'])
        Sn.value = float(data['tensile_strength'])
        Kw.value = float(data['contact_strength'])
        Elastic.value = float(data['elastic_modulus'])
        Puasson.value = float(data['Poisson_ratio'])
        Kmat.value = float(data['worm_material_factor'])
        y_Luis.value = float(data['Lewis_ratio'])

        _ui.messageBox('Данные успешно загружены!')

    except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
 

def exportParameters():
    result = generateData(isForTable = False)
    return result

def getMaterialLibNames():
    materialLibs = _app.materialLibraries
    libNames = []
    for materialLib in materialLibs:
        if materialLib.materials.count > 0:
            libNames.append(materialLib.name)
    return libNames

def getMaterialsFromLib(libName, filterExp):
    global materialsMap
    materialList = None
    if libName in materialsMap:
        materialList = materialsMap[libName]
    else:
        materialLib = _app.materialLibraries.itemByName(libName)
        materials = materialLib.materials
        materialNames = []
        for material in materials:
            materialNames.append(material.name)
        materialsMap[libName] = materialNames
        materialList = materialNames

    if filterExp and len(filterExp) > 0:
        filteredList = []
        for materialName in materialList:
            if materialName.lower().find(filterExp.lower()) >= 0:
                filteredList.append(materialName)
        return filteredList
    else:
        return materialList

def replaceItems(cmdInput, newItems):
    cmdInput.listItems.clear()
    if len(newItems) > 0:
        for item in newItems:
            cmdInput.listItems.add(item, False, '')
        cmdInput.listItems[0].isSelected = True

def getMaterial(materialName):
    materialLibs = _app.materialLibraries
    material = None
    for materialLib in materialLibs:
        materials = materialLib.materials
        try:
            material = materials.itemByName(materialName)
        except:
            pass
        if material:
            break
    return material

def applyMaterialToEntities(material, entities):
    for entity in entities:
        entity.material = material

# Event handler for the validateInputs event.
class GearCommandValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            eventArgs = adsk.core.ValidateInputsEventArgs.cast(args)
        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

