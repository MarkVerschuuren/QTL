
import xlrd

from xlwt import Workbook


wb = Workbook()
Values = wb.add_sheet("Values")


def Openen():
    file = xlrd.open_workbook("Marker Vergelijking ANOVA Test Handmatig.xlsx","r")
    Markers = file.sheet_by_index(0)
    Anthocyanin = file.sheet_by_index(1)

    return Markers, Anthocyanin

def Verwerken_Markers(Markers):
    Marker_dictonary = {}
    for row in range(2,Markers.nrows):
        Marker_dictonary[Markers.cell_value(row,1)] = []
        for column in range(3,Markers.ncols):
            Marker_dictonary[Markers.cell_value(row,1)].append(Markers.cell_value(row,column))

    return Marker_dictonary

def Verwerken_Anthocyanin(Anthocyanin):
    List_Anthocyanin = []
    for row in range(1,Anthocyanin.nrows):
        List_Anthocyanin.append(Anthocyanin.cell_value(row,1))

    return List_Anthocyanin

def Marker_values(Marker, Anthocyanin):
    for key in Marker:
        MarkerA = []
        MarkerB = []
        for index in range(len(Marker[key])):
            if Marker[key][index] == "a":
                if Anthocyanin[index] != "*":
                    MarkerA.append(Anthocyanin[index])
            elif Marker[key][index] == "b":
                if Anthocyanin[index] != "*":
                    MarkerB.append(Anthocyanin[index])
        Anova(MarkerA, MarkerB, key)


def Anova(MarkerA, MarkerB, Locus):
    SStot = 0
    SSbinnen = 0
    F = 3.91 # Waarde gekozen vanuit http://www.ablongman.com/graziano6e/text_site/MATERIAL/Stats/F-tab.pdf. Bij 150 3.91, verschil van 0,01 tussen 25, ongeveer 160 Dif bij onze berekeningen.
    GemA = sum(MarkerA)/len(MarkerA)
    GemB = sum(MarkerB)/len(MarkerB)
    Gtot = (sum(MarkerA)+sum(MarkerB)) / (len(MarkerA) + len(MarkerB))

    for value in MarkerA:
        SSbinnen += (value - GemA)**2
        SStot += (value - Gtot) ** 2
    for value in MarkerB:
        SStot += (value - Gtot)**2
        SSbinnen += (value - GemB)**2

    SStussen = SStot - SSbinnen
    MStussen = SStussen
    MSbinnen = SSbinnen / 160

    P_value = MStussen / MSbinnen
    Values = [Locus, P_value, GemA, GemB, sum(MarkerA), sum(MarkerB), len(MarkerA), len(MarkerB)]

    ExcelPut(Values)


def ExcelPut(Values, row= 0):
    print(row)
    Values.write(0, 0, "Locus")
    Values.write(0, 1, "P_value")
    Values.write(0, 2, "GemA")
    Values.write(0, 3, "GemB")
    Values.write(0, 4, "TotA")
    Values.write(0, 5, "TotB")
    Values.write(0, 6, "#A")
    Values.write(0, 7, "#B")

    for column in range(len(Values)):
        Values.write(row,column,Values[column])
    row += 1



def main():
    Markers, Anthocyanin = Openen()
    Marker_dictonary = Verwerken_Markers(Markers)
    List_Anthocyanin = Verwerken_Anthocyanin(Anthocyanin)
    Marker_values(Marker_dictonary, List_Anthocyanin)
    wb.save("QTL-Values")

main()