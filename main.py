from win32com.client import Dispatch
L = Dispatch('LEAP.LEAPApplication')

L.ActiveArea = "Test Bed"
L.ActiveBranch = "Key Assumptions \ TotalElec"
setscenario = "Baseline"
# Branch = L.Branch
L.ActiveScenario = setscenario

print(L.ActiveBranch.Name)
print(L.ActiveScenario)
print(L.ActiveUnit)
# print(dir(L))
# print(dir(L.Branch))
print(L.Areas.Count)