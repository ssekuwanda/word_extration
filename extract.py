import os
from docx2python import docx2python
import pandas as pd
import datetime

first_name = []
second_name = []
post = []
number = []
city = []
dob = []
email=[]
nationality=[]
passport=[]
university = []
degree = []
spec = []
gpa = []
qualif = []
level = []

# save word documents to a folder called source
for subdir, dirs, files in os.walk(r'.\source'):
    for filename in files:
        filepath = subdir + os.sep + filename

        try:
            if filepath.endswith(".docx") or filepath.endswith(".doc"):
                new = docx2python(filepath)
                one = new.body[1]
                two = new.body[3]
                three = new.body[5]
                # print(three)
                details = []
                for gen in one:
                    for name in gen:
                        for name_one in name: 
                            details.append(name_one)

                sity = []
                for uni in two:
                    for det in uni:
                        for peni in det:
                            sity.append(peni)
                
                prof = []
                for cos in three:
                    for cosi in cos:
                        for course in cosi:
                            prof.append(course)

                first_name.append(details[1])
                second_name.append(details[3])
                post.append(details[5].title())
                number.append(details[9])
                city.append(details[11].title())
                dob.append(details[13])
                email.append(details[15])
                nationality.append(details[17].title())
                passport.append(details[19])
                university.append(sity[1])
                degree.append(sity[3].title())
                spec.append(sity[5])
                gpa.append(sity[11])
                qualif.append(prof[3].title())
                level.append(prof[5])
        except:
            print('error')

df = pd.DataFrame(list(zip(first_name, second_name, post, number, city, dob, email, nationality,passport,
                        university, degree, spec, gpa, qualif, level)), 
               columns =['First name', 'second name', 'post code', 'Phone Number', 'City', 'DOB','Email','Nationality','Passport','university',
                        'degree','Specialization', 'GPA', 'Professional Qualification','Level']) 
df['Name'] = df.apply(lambda x: x['First name'].title() +' '+ x['second name'].title(), axis=1)

# Name of excel file to save to
df.to_excel('Application.xls')