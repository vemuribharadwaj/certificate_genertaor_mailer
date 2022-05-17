import pandas as pd
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw 
from io import BytesIO

df = pd.read_excel('DataSet.xlsx')
#print(df)
#print("Dimention : ",df.shape)
(rows, columns) = df.shape
#print("Number of rows = ",rows)
#print("Number of columns = ",columns)
fileLocation = []
for i in range(rows):
	#Generating Certificate Image
	img = Image.open("BaseImage.jpg")
	draw = ImageDraw.Draw(img)
	font = ImageFont.truetype(font = "c:\\windows\fonts\comic.ttf", size = 30 )
	draw.text((261,527),df.loc[i].NAME,font=font,fill=(0,0,0))
	draw.text((265,681),df.loc[i].COMPETETION_NAME,font=font,fill=(0,0,0))
	draw.text((417,789),df.loc[i].AWARD,font=font,fill=(0,0,0))
	draw.text((629,788),df.loc[i].DAY,font=font,fill=(0,0,0))
	draw.text((969,784),str(df.loc[i].YEAR),font=font,fill=(0,0,0))
	location="certificates\\" + df.loc[i].NAME+df.loc[i].COMPETETION_NAME+".pdf"#in windows "certificates\\"
	img.save( location, "PDF", resolution=100.0)
	fileLocation.append(location)
	print(i," ",df.loc[i].NAME)	
df['CERTIFICATE_LOCAATION'] = fileLocation
df.to_excel("NewDataset.xlsx",sheet_name='participant_certificates_info',index=False)
