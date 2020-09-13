# -*- coding: utf-8 -*-
"""
Created on Mon Jul 13 22:11:50 2020

@author: Vikas Gowda LV
"""

import PIL
from PIL import Image, ImageFont, ImageDraw  
import xlrd 
  
# Give the location of the file 
loc = (r'C:\Users\Vikas\Documents\IEEE\Certificate\Book2.xlsx') 
  
# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  
# For row 0 and column 0 
print(sheet.nrows)
print(sheet.ncols)
for i in range(0,5):
    Name = (sheet.cell_value(i, 0))
    Name = Name.capitalize()
    Name = Name.center(40, " ")
    #print(sheet.cell_value(i, 0))
    
    
    image = Image.open(r'C:\Users\Vikas\Documents\IEEE\Certificate\template.png')  
      
    draw = ImageDraw.Draw(image)  
      
    # specified font size 
    font = ImageFont.truetype(r'C:\Users\Vikas\Documents\IEEE\Certificate\MoongladeDemoBold.ttf',250)  
      
    text = Name
      
    # drawing text size
    print("start")
    draw.text((900,2300),text, fill ="Black", font = font, align ="center")  
    #draw.text((1000,500),text, fill ="Red", font = font, align ="center") 
    print("stop")
   # image.show()
      
    image.save(r'C:\Users\Vikas\Documents\IEEE\Certificate\{}.png'.format(Name))