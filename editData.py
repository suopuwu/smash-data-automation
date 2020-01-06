import openpyxl
sheetName = input('sheet name: ')
wb = openpyxl.load_workbook(sheetName)
if 'GlossaryNotes' in wb.sheetnames:
    del wb['GlossaryNotes']
else:
    print('GlossaryNotes either doesn\'t exist or is named weirdly')
    
for character in wb.sheetnames:
    #if you try to set the title to itself but lowercase, it'll append a 1 because titles are not case sensitive. Thus, we must set a temporary value first.
    wb[character].title = 'temp'
    wb['temp'].title = character.lower()


print('saving...')
wb.save(sheetName)
input('done! press enter to exit.')
