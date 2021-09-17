import ExcelOP

def issameproject(numtotallist,numtemhavegetlist):
    for numtotal in numtotallist:
        # numtotal=numtotal.replace(" ","")
        for numtemhaveget in numtemhavegetlist:
            # numtemhaveget=numtemhaveget.replace(" ","")
            if numtotal==numtemhaveget :
                return True
    return False

excel=ExcelOP.ExcelOP()
totalprojectdata,havegetmoneyprojectdata=excel.getprojectdata()
i=1
for itemtotal in totalprojectdata:
    label=False
    numtotallist=itemtotal[0].split("&")
    # print(numtotallist)
    for itemhaveget in havegetmoneyprojectdata:
        numtemhavegetlist=itemhaveget[0].split("&")
        # print(numtemhavegetlist)
        if issameproject(numtotallist,numtemhavegetlist):
            label=True
            break
    if(not label):
        print(i,",",itemtotal[0],",",itemtotal[1],",",itemtotal[2],",",itemtotal[3])
        i=i+1


