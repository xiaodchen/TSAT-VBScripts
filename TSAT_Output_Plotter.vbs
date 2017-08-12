Set objPlotter = CreateObject("ResultScript.Plotter")
Set objExcel = CreateObject("Excel.Application")

call objPlotter.setOutput("G:\3125\006\40_Ing_Eng\47_Elec\study\ENGINEERING STUDY\DYN STABIL\Plots.docx","docx")
objPlotter.PlotsPerPage = 2 
objPlotter.DoColor = True
objPlotter.DoMark = False

dim BINFILE_1, bus_string, msg, ctg_title, qname, quan_string1, quan_string2, sname, cname
dim xarr, arr, arrcurvelist, qlist, ctgList, scenList
dim a, b, c, d, e, f, x, n, k, R, quan_num, count, min
dim scen_num, ctg_num, num, bus_num, bus_num1, bus_id, bus_id1

BINFILE_1 = "G:\3125\006\40_Ing_Eng\47_Elec\study\ENGINEERING STUDY\DYN STABIL\Schuler_2019_Winter_Peak.bin" 

Call objPlotter.XAxisRange(0.000000,5.000000)
'Y axis range is not defined specifically but based on curves 
'Call objPlotter.YAxisRange(0.000000,1.500000)

Set reader = CreateObject("ResultScript.BinReader")

reader.file = BINFILE_1

scenList = reader.scenList()

'all these List starts at 1 and not 0, ctgList, qList, scenList, and arrcurveList
For scen_num = 1 to UBound(scenList)
	
	'specify scenario 
	reader.scen = scen_num
	
	'specify scenario name 
	sname = scenList(scen_num)
	
	ctgList = reader.ctgList()
	
	'plot all contingencies 
	For ctg_num = 1 to UBound(ctgList)
	
		'specify contingency name 
		cname = ctgList(ctg_num)
		
		qList = reader.quanList()
		
		'plot all monitored quantities for each specified contingency 
		For quan_num = 1 to UBound(qList)
			
			'parse for quantity type and name 
			quan_string = qList(quan_num)
			e = Split (quan_string,":", 2)
			quan_string1 = e(0)
			quan_string2 = e(1)
			
			'specify monitored quantity type and name 
			reader.quan = quan_string1
			qname = quan_string2
			
			'specify contingency 
			reader.ctg = ctg_num
			
			'return list of graphs 
			arrcurvelist = reader.curveList()
			
			'plot all curves or bus monitored for each quantity and contingency,
			For i = 1 to UBound(arrcurvelist) 
				
				n = UBound(arrcurvelist)-i
				
				min = objExcel.Application.WorksheetFunction.Min(5, n)
				
				'plot 6 curves per graph 
				For count = 0 to min
					x = i + count
					
					'Msgbox "min, count, x, i, n, total = " & min & "," & count & "," & x & "," & i & "," & n & "," & UBound(arrcurvelist)
					
					'parse for bus number & bus id; format is '00015:       8 [BUS6    13.813.8]  1 
					bus_string = arrcurvelist(x)
					a = Split (bus_string, "[", 2)
					b = Split (a(0), ":", 2)
					bus_num1 = b(1)
					bus_num = CLng(bus_num1)			
					R = Vartype(bus_num)
					bus_id = Right(bus_string, 2)
					
					'contingency title 
					ctg = reader.ctgList()
					
					'specify bus number, bus ID
					reader.bus1 = bus_num
					reader.id = bus_id
			
					'Msgbox "scen, ctg, quantity, bus: " & scen_num & "," & ctg_num & "," & quan_num & "," & bus_num & "," & bus_id
					
					xarr = reader.timeValues()	
					arr = reader.curveValues()
				
					'Create curve 
					Call objPlotter.AddTYCurve("Bus No.: " & bus_num, xarr, arr)	
				
				Next
				
				'Plot all curves created so far
				Call objPlotter.DoPlot("Scen. No. " & scen_num & ": " & sname & " - Ctg. No. " & ctg_num &" - " & qname, "Time (sec)","", qname)	
				
				'counter needs to increase by the number of curves plotted so far
				i = i + min

			Next
			
		Next	

	Next 
	
Next 

Msgbox "All Finished"

objPlotter.Finish()
objExcel.Quit

Set objPlotter = Nothing
Set objExcel = Nothing