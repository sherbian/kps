import os
import re
import math
import random
import subprocess
import pythoncom

from win32com.client import Dispatch, gencache

TEXT_ITEM_ARR = 4

def is_running():
    proc_list = \
    subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"', shell=False, stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False

def get_compas():
    kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))


    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

    pythoncom.CoInitialize()

    return kompas6_constants, kompas6_constants_3d, kompas6_api5_module, kompas_object, kompas_api7_module, application

print('Running:', is_running())

kompas6_constants, kompas6_constants_3d, kompas6_api5_module, kompas_object, kompas_api7_module, application = get_compas()
Documents = application.Documents

kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentDrawing, True)

kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()




class kompasText:
    def __init__(self, x=39, y=130,style=0) -> None:
        self.style = style
        iParagraphParam = kompas6_api5_module.ksParagraphParam(kompas_object.GetParamStruct(kompas6_constants.ko_ParagraphParam))
        iParagraphParam.Init()
        iParagraphParam.x = x
        iParagraphParam.y = y
        iParagraphParam.ang = 0
        iParagraphParam.height = 4.219972133636
        iParagraphParam.width = 3.518958091736
        iParagraphParam.hFormat = 0
        iParagraphParam.vFormat = 0
        iParagraphParam.style = 1
        iDocument2D.ksParagraph(iParagraphParam)
    def addText(self, text):
        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 1
        iTextItemArray = kompas_object.GetDynamicArray(TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = text
        iTextItemParam.type = 0
        iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "GOST type A"
        iTextItemFont.height = 5
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)
        iDocument2D.ksTextLine(iTextLineParam)
    def end(self):
        obj = iDocument2D.ksEndObj()
        iDrawingText = kompas_object.TransferReference(obj, 0)
        iDrawingText.Allocation = self.style
        kompas_api7_module.IDrawingObject(iDrawingText).Update()

class kompasQuad:
    def __init__(self, x, y) -> None:
        self.x = x
        self.y = y
    def draw(self, x, y, col):
        boundaries = iDocument2D.ksNewGroup(1)
        iDocument2D.ksContour(1)
        iDocument2D.ksNewGroup(1)
        obj = iDocument2D.ksLineSeg(x, y, x+self.x, y, 1)
        obj = iDocument2D.ksLineSeg(x+self.x, y, x+self.x, y+self.y, 1)
        obj = iDocument2D.ksLineSeg(x+self.x, y+self.y, x, y+self.y, 1)
        obj = iDocument2D.ksLineSeg(x, y+self.y, x, y, 1)
        iDocument2D.ksEndGroup()
        obj = iDocument2D.ksEndObj()

        iDocument2D.ksEndGroup()
        obj = iDocument2D.ksColouringEx(col, boundaries)
        iColouring = kompas_object.TransferReference(obj,  0)


class DiagrammWheel:
    radius = 60
    formats = {
        'A4': {'x':210, 'y':297},
        'A3': {'x':297, 'y':420},
        'A2': {'x':420, 'y':594},
        'A1': {'x':594, 'y':841},
        'A0': {'x':841, 'y':1189}
    }
    class pieceOfWheel:
        def __init__(self) -> None:
            self.data = []
        def add(self,*, name, proc=0,value, x=0, y=0, color='rand'):
            self.data.append({'n':name, 'p':proc,'x':x, 'y':y,'c':color, 'v':value, 'a':0})
        def all(self):
            # return sum(map(lambda v: v.get('v'), self.data))
            return 35324950.9
        def get(self) -> list:
            a = self.all()
            angle = 0
            for v in self.data:
                proc = v['v'] / a
                angle += 360 * proc
                v['p'] = proc
                v['a'] = angle
            return self.data


    def __init__(self,*, format = 'A4', offsetAngle = 0, procOffset=10) -> None:
        self.data = self.pieceOfWheel()
        self.angleOffset = offsetAngle
        self.procOffset = procOffset
        self.quad = kompasQuad(3, 3)
        if format in self.formats.keys():
            self.format = format
        else:
            raise Exception("Sorry, format not found")
        self.mathStartPoints()

    
    def mathStartPoints(self):
        self.x = self.formats[self.format.upper()].get('x')/2 + 12.5
        self.y = self.formats[self.format.upper()].get('y')-45-self.radius
        text = kompasText(self.x, self.y + self.radius + 20, style=1)
        text.addText('Диаграмма затрат на производство детали "Корпус"')
        text.end()

    def normalize(self):
        normalize = lambda a: a % 360 if a > 0 else -a % 360
        last_angle = 0
        data = self.data.get()
        self.angleOffset = -(data[0]['p']*0.5 * 360)

        for v in data:
            v['x'] = self.x + self.radius * math.cos(math.radians(-normalize(v.get('a') + self.angleOffset)))
            v['y'] = self.y + self.radius * math.sin(math.radians(-normalize(v.get('a') + self.angleOffset)))
            
            a = last_angle + (v.get('p')*0.5 * 360)

            v['cx'] = self.x + (self.radius+self.procOffset) * math.cos(math.radians(-normalize(a + self.angleOffset)))
            v['cy'] = self.y + (self.radius+self.procOffset) * math.sin(math.radians(-normalize(a + self.angleOffset)))

            last_angle = v.get('a')
            v['c'] = int('0x{:02X}{:02X}{:02X}'.format(random.randint(0, 256),random.randint(0, 256),random.randint(0, 256)), 16)
            

    def drawing(self):
        self.normalize()
        ''' RUNNING '''
        last_x = self.x + self.radius * math.cos(math.radians(-(lambda a: a % 360 if a > 0 else -a % 360)(self.angleOffset)))
        last_y = self.y + self.radius * math.sin(math.radians(-(lambda a: a % 360 if a > 0 else -a % 360)(self.angleOffset)))

        pos_y = 110

        for v in self.data.get():
            boundaries = iDocument2D.ksNewGroup(1)
            iDocument2D.ksContour(1)
            iDocument2D.ksNewGroup(1)
            obj = iDocument2D.ksArcByPoint(self.x, self.y, self.radius, last_x, last_y, v.get('x'), v.get('y'), -1, 1)
            obj = iDocument2D.ksLineSeg(self.x, self.y, v.get('x'), v.get('y'), 1)
            obj = iDocument2D.ksLineSeg(self.x, self.y, last_x, last_y, 1)
            iDocument2D.ksEndGroup()
            obj = iDocument2D.ksEndObj()

            iDocument2D.ksEndGroup()         # b  g  a
            obj = iDocument2D.ksColouringEx(v.get('c'), boundaries)

            text = kompasText(v.get('cx'), v.get('cy'), style=1)
            p = v.get('p')* 100
            text.addText(f'{p:.1f}%')
            text.end()
            if v.get('n') == 'Отчисления на страховые взносы ФОТ основных производственных рабочих':
                self.quad.draw(30, pos_y, v.get('c'))
                text = kompasText(35, pos_y-0.5, style=0)
                text.addText('Отчисления на страховые взносы ФОТ основных')
                text.end()
                pos_y -= 6
                text = kompasText(35, pos_y-0.5, style=0)
                text.addText('  производственных рабочих' + f' - {(v.get("p")*100):.1f}%')
                text.end()

            else:
                self.quad.draw(30, pos_y, v.get('c'))
                text = kompasText(35, pos_y-0.5, style=0)
                text.addText(v.get('n') + f' - {(v.get("p")*100):.1f}%')
                text.end()

            pos_y -= 6
            last_x = v.get('x')
            last_y = v.get('y')
        

wheel = DiagrammWheel(offsetAngle=False)

wheel.data.add(name='Сырье и материалы', value=10416000)
wheel.data.add(name='Основная зарплата основных производственных рабочих', value=4019640)
wheel.data.add(name='Дополнительная зарплата дополнительных основных рабочих', value=401964)
wheel.data.add(name='Отчисления на страховые взносы ФОТ основных производственных рабочих', value=1335324.4)
wheel.data.add(name='Общепроизводственные расходы', value=8442020.2)
wheel.data.add(name='Общехозяйственные расходы', value=6102400)
wheel.data.add(name='Комерческие расходы', value=4607602.3)



wheel.drawing()














