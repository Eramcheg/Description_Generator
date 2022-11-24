import Config
import openai
import openpyxl
openai.api_key= Config.api_key

def test(text):
    response = openai.Completion.create(
        model="text-davinci-002",
        prompt=text,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    return response.choices[0].text

#Description Generator



location=('G:\FIles\Materials1.xlsx')
workbook=openpyxl.load_workbook(location)
worksheet= workbook.get_sheet_by_name("1")
range_start=2
range_finish=21
for i in range (range_start,range_finish):
    KEY_WORDS_SAMPLE=""

    Color=worksheet['O' + str(i)].value

    Material=worksheet['K'+str(i)].value

    Product_Name=str(worksheet['H'+str(i)].value)

    Plating=worksheet['L' + str(i)].value

    Stone_Type=worksheet['N' + str(i)].value

    Item_Group=worksheet['I' + str(i)].value

    # Appending Color to Sample
    if Color != None:
        KEY_WORDS_SAMPLE +="Color:"+str(Color)  +"\n"


    #Appending Material Type to Sample
    if Material!=None:
        if Material == "Steel":
            KEY_WORDS_SAMPLE+="Material: Stainless steel\n"
        else:
            KEY_WORDS_SAMPLE+="Material: "+str(Material) +"\n"


    #Appending Product Name to Sample
    KEY_WORDS_SAMPLE+= "Product: "+Product_Name + "\n"


    #Appending Plating Type to Sample
    if Plating != None:
        if (str(Plating)).upper() == "RH":
            KEY_WORDS_SAMPLE += "Plating: Rhodium plating\n"
        if (str(Plating)).upper() == "RG":
            KEY_WORDS_SAMPLE += "Plating: Rose Gold plating\n "
        if (str(Plating)).upper() == "G":
            KEY_WORDS_SAMPLE += "Plating: Gold plating\n"


    #Appending Stone Type to Sample
    if Stone_Type != None:
        if (str(Stone_Type)).find(',')!=-1:
            KEY_WORDS_SAMPLE += "Types of jewelry stone: "+Stone_Type+"\n"
        elif (str(Stone_Type)).find(',')==-1:
            KEY_WORDS_SAMPLE += "Jewelry stone type:" + Stone_Type + "\n"

    KEY_WORDS_SAMPLE+="Write a description for this product\n"


    res=test(KEY_WORDS_SAMPLE)
    print(res)

    worksheet['U'+str(i)] = res                     #Appending result to U[i] column
    workbook.save('G:\FIles\Materials1.xlsx')       # Saving updated workbook
