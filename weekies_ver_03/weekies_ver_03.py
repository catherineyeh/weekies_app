import Members
import xlsxwriter 
import module1
import random
from openpyxl import load_workbook
from xlutils import copy

file = load_workbook(r'C:\proj\weekies_ver_03\weekies_ver_03\weekies.xlsx')
sheet = file.active

#whole team
team = module1.members(sheet)

candidates = module1.getCandidates(team)
#department candidates
po = 'Product Operations'
dev = 'Developers'
design = 'Designers'
departments = []
poTeam = module1.depMembers(candidates, po)

devTeam = module1.depMembers(candidates, dev)

designTeam = module1.depMembers(candidates, design)
#select two departments
#select the ones that are not empty
if len(poTeam):
    departments.append(po)
if len(devTeam):
    departments.append(dev)
if len(designTeam):
    departments.append(design)

random.seed(a = None)

round2Candidates = []
if len(poTeam):
    candidateOne = random.choice(poTeam)
    round2Candidates.append(candidateOne)
if len(devTeam):
    candidateTwo = random.choice(devTeam)
    round2Candidates .append(candidateTwo)
if len(designTeam):
    candidateThree = random.choice(designTeam)
    round2Candidates .append(candidateThree)

indices = random.sample(range(0, len(round2Candidates)), 2)
firstIndex = indices[0]
secIndex = indices[1]
weekieOne = round2Candidates[firstIndex]
weekieTwo = round2Candidates[secIndex]

print("This week's weekies are:")
print(weekieOne.name + " and " + weekieTwo.name)
print("\nSee you next week!")

module1.updateInfoInFile(weekieOne, weekieTwo, sheet)

file.save('weekies.xlsx')

