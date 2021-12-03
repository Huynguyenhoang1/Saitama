# -*- coding: utf-8 -*-
__title__ = 'Duplicate\nSheet'
from Autodesk.Revit import DB 
import sys, os  
from pyrevit import (forms, revit)  
from pyrevit import script

import clr 
import pyrevit 

doc = __revit__.ActiveUIDocument.Document 
uidoc = __revit__.ActiveUIDocument 
uiapp = __revit__ 
app = uiapp.Application 


sel = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]
if sel:
	try:
		sheet= [i for i in sel if i.ViewType.Equals(DB.ViewType.DrawingSheet)][0]
		x = sheet.LookupParameter('Sheet Name').AsString()
		#prosset sheet name
		result = ''.join([i for i in x if not i.isdigit()])
		a = x[-1:]
		t = str(a.isdigit())
		g = result
		with revit.Transaction("dublicate Sheet"):

			if t == 'False':#if not number have a sheet name
				o = x + '01'
				sheet.LookupParameter('Sheet Name').Set(o)
				g = result+'02'

		if t == 'True':# if number have a sheet name
			p = int(a)+1
			g = result+'0'+str(p)#set parameter for g


		#prosset sheet number
		y = sheet.LookupParameter('Sheet Number').AsString()
		n = int(y[-1:])+1
		m = y[:-1]+str(n)
		#proset シート 発行目的
		a1 = sheet.LookupParameter('シート 発行目的').AsString()
		
	except: pass
elif uidoc.ActiveView.ViewType.Equals(DB.ViewType.DrawingSheet):
	sheet =uidoc.ActiveView
else: 
	forms.alert("No sheet selected", ok=True)
	sys.exit()

vportIdlist= sheet.GetAllViewports() #returns elementIds

viewtypelist = [DB.ViewType.Legend]

# listcomp: tuple (3elem: view, XYZ-Point, vportTypeId)
view_xyz_vptypeId= [(doc.GetElement(doc.GetElement(vpid).ViewId),
							doc.GetElement(vpid).GetBoxCenter(), 
							doc.GetElement(vpid).GetTypeId())
					for vpid in vportIdlist if doc.GetElement(doc.GetElement(vpid).ViewId).ViewType in viewtypelist]


# get TitleBlock of Sheet to dublicate
tibl = DB.FilteredElementCollector(doc).OwnedByView(sheet.Id).OfCategory(DB.BuiltInCategory.OST_TitleBlocks) \
			.ToElements()

#Get the Type of the TitleBlock Instance 
tibltype = tibl[0].Symbol 
tiblXYZ = tibl[0].Location.Point 

with revit.Transaction("dublicate Sheet"):
	sheetA = DB.ViewSheet.Create(doc, tibltype.Id)

	#Place 2nd TitleBlock, if it exists
	if len(tibl) >1:
		doc.Create.NewFamilyInstance(DB.XYZ(0,0,0), tibl[1].Symbol, sheetA)
	elif len(tibl) >2: print "MORE Than 2 TitleBlocks on Sheet"
	
	#create legend viewports
	for view, xyzpt, vptypeid in view_xyz_vptypeId:
		vp= DB.Viewport.Create(doc, sheetA.Id, view.Id, xyzpt.Subtract(tiblXYZ))
		
		# Match the Type of the Viewport with ChangeTypeId-Method
		vp.ChangeTypeId(vptypeid)
	h = sheetA.LookupParameter('Sheet Name').Set(g)
	k =  sheetA.LookupParameter('Sheet Number').Set(m)
	j =  sheetA.LookupParameter('シート 発行目的').Set(a1)
