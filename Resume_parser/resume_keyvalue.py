from flask import Flask,  request, jsonify
from werkzeug.utils import secure_filename
import os
import re
import aspose.words as aw
import pathlib
import spacy
from indian_cities.dj_city import cities
from nltk import flatten, word_tokenize
import slate3k as slate
import win32com.client

app = Flask(__name__)

UPLOAD_FOLDER = r'C:\My_Work\Akash_resume_parser\uploads'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
index = 0

ALLOWED_EXTENSIONS = {'pdf', 'docx'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return 'Hello welcome to login site. Try api'


@app.route('/Form', methods=['POST'])
def Form():
    global index

    if 'files[]' not in request.files:
        resp = jsonify({'message': 'No file part in the request'})
        resp.status_code = 400
        return resp

    files = request.files.getlist('files[]')
    errors = {}
    success = False

    for file in files:
        print(f"type of : {type(file)}")
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            success = True
        else:
            errors[file.filename] = 'File type is not allowed'

    if success and errors:
        errors['message'] = 'File(s) successfully uploaded'
        resp = jsonify(errors)
        resp.status_code = 500
        return resp
    if success:
        resp = jsonify({'message': 'Files successfully uploaded'})
        resp.status_code = 201
        files = filename
        res_model = r'C:\My_Work\Akash_resume_parser\uploads\\'
        resume = res_model + files
        x = pathlib.PurePosixPath(resume).suffix

        if x == ".pdf":
            text = ""
            with open(resume, 'rb') as f:
                text_c = slate.PDF(f)
            for word in text_c:
                text += word

            text = os.linesep.join([s for s in text.splitlines() if s])
            text=re.sub("\r"," ",text)

        elif x == ".docx":
            doc = aw.Document(resume)
            resume = resume.replace('.docx', '.pdf')
            # Save as PDF
            doc.save(resume)
            text = ""
            with open(resume, 'rb') as f:
                text_c = slate.PDF(f)
            for word in text_c:
                text += word
            text = re.sub("•", "", text)
            text = os.linesep.join([s for s in text.splitlines() if s])
            text = re.sub("\r", " ", text)
            text = re.sub("Evaluation Only. Created with Aspose.Words. Copyright 2003-2023 Aspose Pty Ltd.", "", text)
            text = re.sub("Created with an evaluation copy of Aspose.Words. To discover the full versions of our APIs",
                          "", text)
            text = re.sub("please visit: https://products.aspose.com/words/", "", text)
            text = re.sub("visit: https://products.aspose.com/words/", "", text)
            text = re.sub("please", "", text)



        else:
            text = None
            print("incorrect file extension")

        def parsing(text, Keywords, a):

            Keywords=list(Keywords)
            if a==0 :
                text = text.replace("\n", " ")
                text = text.replace("[^a-zA-Z0-9]", " ");
                re.sub('\W+', '', text)

            elif a == 1:
                text = text.replace("\n", " ")
                text = text.replace("[^a-zA-Z0-9]", " ");
                re.sub('\W+', '', text)
                text = text.lower()

            elif a == 2:
                split_text = text.split("\n")
                text_list = []
                for i in split_text:
                    cap_text = i.capitalize()
                    string = re.sub("/\d\.\s+|[a-z]\)\s+|•\s+|[A-Z]\.\s+|[IVX]+\.\s+/g", "", cap_text);
                    n = string.split("\n")
                    text_list.extend(n)
                new_list = [x for x in text_list if x != ' ']
                old_list = [x for x in new_list if x != '  ']
                text = " ".join(old_list)

            parsed_content = {}
            content = {}
            indices = []
            keys = []
            for key in Keywords:
                try:
                    content[key] = text[text.index(key) + len(key):]
                    indices.append(text.index(key))
                    keys.append(key)
                except:
                    pass

            if len(keys)!=0:
                zipped_lists = zip(indices, keys)
                sorted_pairs = sorted(zipped_lists)

                tuples = zip(*sorted_pairs)
                indices, keys = [list(tuple) for tuple in tuples]

                # Keeping the required content and removing the redundant part
                content = []
                for idx in range(len(indices)):
                    if idx != len(indices) - 1:
                        content.append(text[indices[idx]: indices[idx + 1]])
                    else:
                        content.append(text[indices[idx]:])
                for i in range(len(indices)):
                    parsed_content[keys[i]] = content[i]
            return parsed_content



        # To extract Email
        def get_email_addresses(string):
            r = re.compile(r'[\w\.-]+@[\w\.-]+')
            return r.findall(string)

        # To extract Phone number
        def get_phone_numbers(string):
            r = re.compile(r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})')
            phone_numbers = r.findall(string)
            return [re.sub(r'\D', '', num) for num in phone_numbers]

        nlp = spacy.load('en_core_web_sm')

        # To extract Location
        def extract_loc(texts):
            nlp_text = nlp(texts)
            tokens = [token.text for token in nlp_text if not token.is_stop]
            places = []
            city_name = flatten(cities)
            city = list(map(str.lower, city_name))
            for words in tokens:
                if (words.lower() in city):
                    places.append(words.title())
            places = list(set(places))
            return places

        def extract_skills(texts):
            nlp_text = nlp(texts)
            tokens = [token.text for token in nlp_text if not token.is_stop]
            skills = ["machine learning", "deep learning", "nlp", "natural language processing", "mysql", "sql",
                      "django", "react","js",
                      "ethical hacking",
                      "tensorflow", "opencv", "mongodb", "artificial intelligence", "docker", "pyspark", "ccna",
                      "kubernetes", "python", "c++",
                      "terraform", "css", "azure", "github", "aws", "tableau"]
            skillset = []
            for token in tokens:
                if token.lower() in skills:
                    skillset.append(token.lower())

            # check for bi-grams and tri-grams (example: machine learning)
            skillset = [*set(skillset)]
            return (skillset)

        def extract_secondaryskills(texts):
            nlp_text = nlp(texts)
            tokens = [token.text for token in nlp_text if not token.is_stop]
            skills = ["teamwork", "mentoring", "leadership", "building client relationships",
                      "communication skills", "adaptability", "problem solving", "creativity", "work ethic",
                      "interpersonal skills", "time management", "listening skills", "social etiquette",
                      "management skills", "organizational skills", "analytical skills", "collaboration",
                      "critical thinking", "presentation skills", "positive attitude", "assertiveness", "empathy"]
            sskillset = []
            for token in tokens:
                if token.lower() in skills:
                    sskillset.append(token.lower())

            sskillset = [*set(sskillset)]
            return sskillset

        def passed_ot(values):
            z = re.sub(r"[\([{})\]]", "", values)
            y = re.sub('[a-z]+', '', z)
            pattern = re.compile('[2][0]\d\d')
            matches = pattern.findall(y)
            strtolst = re.split("[\s-]",y)
            out = any(check in matches for check in strtolst)
            if out == True:
                passed = max(matches)
                return(passed)
            else:
                a = []
                for i in strtolst:
                    pattern_2 = re.compile(r'^20....\d$')
                    a.append(pattern_2.findall(i))
                res = [ele for ele in a if ele != []]
                for x in res:
                    out_2 = any(check in x for check in strtolst)
                if out_2 == True:
                    passed = 0
                    b = []
                    for match in x:
                        e = match.split("-")
                        for i in e:
                            b.append(i)
                            if len(i) < 4:
                                i = "20" + i
                                b.append(i)
                    for x in b:
                        if len(x) < 4:
                            b.remove(x)
                            passed = max(b)
                    return(passed)

        def extract_bach_deg(texts):
            li = []
            deg = "Not mentioned"

            for item in texts.split("  "):
                li.append(item)

            bachelor_list = ["bachelor", "bca", "b.c.a", "b.e", "b.a", "bsc", "b.com", "b.ed", "btech", "b.tech",
                             "b.f.tech",
                             "be"]
            for word in li:
                for temp in bachelor_list:
                    if temp in word.lower():
                        deg = word

            return deg

        def extract_bach_coll(texts):
            li = []
            lis = []
            for item in texts.split("  "):
                li.append(item)
            for word in li:
                if "\t" in word:
                    word = re.sub("\t", "", word)
                    lis.append(word)
                else:
                    lis.append(word)
            lis = [i for i in lis if i]
            bachelor_list = ["bachelor", "bca", "b.c.a", "b.e", "b.a", "bsc", "b.com", "b.ed", "btech", "b.tech",
                             "b.f.tech",
                             "be"]

            for word in lis:
                for temp in bachelor_list:
                    if temp in word.lower():
                        index = lis.index(word)
            college = ["college", "university", "institute"]

            def index_exist(name, index):
                try:
                    c = (name[index])
                    return True
                except:
                    return False

            for colle in college:
                if index_exist(lis, index) and colle in lis[index].lower():
                    coll = lis[index]
                elif index_exist(lis, index - 1) and colle in lis[index - 1].lower():
                    coll = lis[index - 1]
                elif index_exist(lis, index - 2) and colle in lis[index - 2].lower():
                    coll = lis[index - 2]
                elif index_exist(lis, index - 3) and colle in lis[index - 3].lower():
                    coll = lis[index - 3]
                elif index_exist(lis, index + 1) and colle in lis[index + 1].lower():
                    coll = lis[index + 1]
                elif index_exist(lis, index + 2) and colle in lis[index + 2].lower():
                    coll = lis[index + 2]
                elif index_exist(lis, index + 3) and colle in lis[index + 3].lower():
                    coll = lis[index + 3]

            coll = re.sub(r"\s\s+", " - ", coll)
            return coll

        def extract_mast_deg(texts):
            li = []
            for item in texts.split(" "):
                li.append(item)

            master_list = ["master", "mca", "pg", "mba", "m.c.a", "m.e", "m.a", "msc", "m.ed", "med", "mtech", "m.tech",
                           "m.f.tech", "me"]

            for i in li:
                if i != "" and i != " ":
                    lis = re.split("\s", i)
                    lis = set(lis)
                    # stops when 1st match element is found
                    deg = next((ele for ele in master_list if ele in lis), None)
                    if deg is not None:
                        break

            return deg

        def extract_mast_coll(texts):
            li = []
            lis = []

            for item in texts.split("  "):
                li.append(item)
            for word in li:
                if "\t" in word:
                    word = re.sub("\t", "", word)
                    lis.append(word)
                else:
                    lis.append(word)
            lis = [i for i in lis if i]
            master_list = ["master", "mca", "pg", "mba", "m.c.a", "m.e", "m.a", "msc", "m.ed", "med", "mtech", "m.tech",
                           "m.f.tech", "me"]
            for word in lis:
                for temp in master_list:
                    if temp in word.lower():
                        index = lis.index(word)

            def index_exist(name, ind):
                try:
                    c = (name[ind])
                    return True
                except:
                    return False

            college = ["college", 'university', "institute"]
            for colle in college:
                if index_exist(lis, index) and colle in lis[index].lower():
                    coll = lis[index]
                elif index_exist(lis, index - 1) and colle in lis[index - 1].lower():
                    coll = lis[index - 1]
                elif index_exist(lis, index - 2) and colle in lis[index - 2].lower():
                    coll = lis[index - 2]
                elif index_exist(lis, index - 3) and colle in lis[index - 3].lower():
                    coll = lis[index - 3]
                elif index_exist(lis, index + 1) and colle in lis[index + 1].lower():
                    coll = lis[index + 1]
                elif index_exist(lis, index + 2) and colle in lis[index + 2].lower():
                    coll = lis[index + 2]
                elif index_exist(lis, index + 3) and colle in lis[index + 3].lower():
                    coll = lis[index + 3]
            return coll

        def extract_educ(texts):
            total_education = re.sub("•", " ", texts)
            total_education = re.split("  ", total_education)
            del total_education[0]
            while "" in total_education:
                total_education.remove("")
            education = {"Education": total_education}
            return education

        Keywords = ["education", "projects profile", "certification", "summary", "Projects", "accomplishments",
                    "executive profile", "professional profile","objective","edu",
                    "personal profile", "work background", "academic profile", "other activities", "qualifications",
                    "experience", "interests", "skills",
                    "achievements", "publications", "work experience", "technical skills", "publication","professional experience",
                    "certifications", "workshops", "projects",
                    "internships", "professional summary","Professional Experience", "trainings", "project work", "personal info",
                    "personal details", "DECLARATION", "hobbies",
                    "overview", "objective", "position of responsibility", "jobs", "Roles and responsibilities",
                    "Roles & Responsibilities",
                    "Roles and Responsibilities", "Work Experience", "Technical skills"]
        x = lambda a=text, b=tuple(Keywords), c=1: parsing(a, b, c)
        parsed_content = x()
        parsed_content = {k: v for k, v in parsed_content.items() if v}
        # To print Location
        loc = extract_loc(text)

        # To extracting Phone_number
        phone_number = get_phone_numbers(text)
        if len(phone_number) <= 10:
            parsed_content['Phone number'] = phone_number

        # To get Email
        email = get_email_addresses(text)
        parsed_content['E-mail'] = email

        # To extract designation
        designation1 = ["lead", "developer", "trainee", "engineer",
                        "designer", "Manager"]
        designation = []
        for key, value in parsed_content.items():
            if key == "designation":
                designation.append(value)
            else:
                m = word_tokenize(text)
                lst1 = []
                for i in designation1:
                    if i in m:
                        lst1.append(m[m.index(i) - 1] + " " + m[m.index(i)])
                tuple1 = tuple(lst1)
        designation = list(tuple1)

        # To print Skills
        skill = extract_skills(text)
        if len(skill) ==0:
            skill = "None"
        # To print Secondary Skills
        sskill = extract_secondaryskills(text)
        if len(sskill) ==0:
            sskill = "None"
        # Passed out year
        try:
            for keys, values in parsed_content.items():
                if (keys == "academic profile") or keys == 'edu' or (keys == "education") :
                    passed = passed_ot(values)
            passed_out = passed
        except:
            passed_out = "Not mentioned"

        # To get projects
        project=None
        Keywords_s = ["Summary", "Accomplishments", "Executive profile", "Professional profile",
                    "Personal profile", "Work background", "Academic profile",
                    "Other activities", "Qualifications", "Professional summary", "Professional experience",
                    "Interests", "Experience summary", "Declaration",
                    "Executive summary", "Work experience", "Key skills", "Technical skills",
                    "Technical proficiency ",
                    "Education", "Skills", "Achievements", "Publications",
                    "Certifications", "Workshops", "Project details",
                    "Projects", "Internships", "Trainings", "Hobbies", "Overview", "Mail",
                    "Objective", "Position of responsibility", "Jobs", "Summary", "Work experience",
                    "Technical skills",
                    "Project experience", "Project", "Project Summary",
                    "Personal Details", "Project Experience", "Project", "Mini project", "Academic project",
                    "Personal details ",
                    "Project experience", "College projects", "mini projects", "final projects", "Final rojects"]
        projects = lambda a=text, b=tuple(Keywords_s), c=2: parsing(a, b, c)
        parsed_content_project = projects()
        parsed_content = {k: v for k, v in parsed_content.items() if v}
        for k, v in parsed_content_project.items():

            if (k == "Projects") or (k == "Project details") or (k == "Project Experience") or \
                    (k == "Project summary") or (k == "Project") or (k == "Mini project") \
                    or (k == "Academic project") or (k == "Project experience") or (k == "College projects") \
                    or (k == "mini projects") or (k == "final projects") or (k == "Final projects"):
                project = v

        # To get Bachelor education
        try:
            for keys, values in parsed_content.items():
                if keys == "education" or keys == 'edu':
                    bach_deg = extract_bach_deg(values)
            bachelor_degree = bach_deg
        except:
            bach_deg = "Not mentioned"
            bachelor_degree = bach_deg

        # To get Bachelor college
        try:
            for keys, values in parsed_content.items():
                if keys == "education"or keys == 'edu':
                    bach_coll = extract_bach_coll(values)
            bachelor_college = bach_coll
        except:
            bach_coll = "Not mentioned"
            bachelor_college = bach_coll

        # To get Master education
        x = lambda a=text, b=tuple(Keywords), c=1: parsing(a, b, c)
        parsed_content = x()
        try:
            for keys, values in parsed_content.items():
                if keys == "education" :
                    mast_deg = extract_mast_deg(values)

            mast_deg = mast_deg
        except:
            mast_deg = "Not mentioned"
            mast_deg = mast_deg

        # To get Master college
        try:
            if mast_deg != "Not mentioned" :
                for keys, values in parsed_content.items():
                    if keys == "education" :
                        mast_college = extract_mast_coll(values)
                mast_college = mast_college
            else : mast_college = "Not mentioned"
        except:
            mast_college = "Not mentioned"
            mast_college = mast_college

        # To extract qualification
        qualify = "None"
        print("=========")
        print(mast_deg)
        if mast_deg != "Not mentioned" and mast_deg:
            master_list = ["master", "mca", "pg", "mba", "m.c.a", "m.e", "m.a", "msc", "m.ed", "med", "mtech",
                           "m.tech",
                           "m.f.tech", "me"]
            for word in master_list:
                if word in mast_deg:
                    if word == "master" or "me":
                        word = "master of engineering"
                        qualify = word.title()
                    else:
                        qualify = word.title()
        elif bach_deg != "Not mentioned":
            bachelor_list = ["bachelor", "bca", "b.c.a", "b.e", "b.a", "bsc", "b.com", "b.ed", "btech", "b.tech",
                             "b.f.tech",
                             "be"]
            for word in bachelor_list:
                if word in bach_deg:
                    if word == "bachelor":
                        word = "bachelor of engineering"
                        qualify = word.title()
                    else:
                        qualify = word.title()
        else:
            qualify = "Not Found"
        qualifcation = qualify

        # To get  education

        try:
            for keys, values in parsed_content.items():
                if keys == "education" or keys == 'edu':
                    educate = extract_educ(values)
            educ = educate
        except:
            educ = None
        education = educ

        # To get professional summary
        Keywords_p = ["Education", "Accomplishments", "Executive Profile", "Professional Profile",
                    "Personal Profile", "Work Background", "Roles and Responsibilities", "Academic Profile",
                    "Other Activities", "Qualifications", "Professional Summary", "Professional Experience",
                    "Interests", "Experience Summary", "EXPERIENCE SUMMARY", "PROFESSIONAL SUMMARY",
                    "EXECUTIVE SUMMARY", "WORK EXPERIENCE", "KEY SKILLS", "TECHNICAL SKILLS", "Technical Proficiency ",
                    "Education", "EDUCATION", "Skills", "personal info", "Achievements", "Publications", "Publication",
                    "EXTRA CURRICULAR ACTIVITIES", "Certifications", "Workshops", "Projects", "Internships",
                    "Trainings","OBJECTIVE",
                    "Hobbies", "Overview",
                    "Objective", "Position of Responsibility", "Jobs", "SUMMARY", "Work Experience", "Technical Skills"
                    ]

        x = lambda a=text, b=tuple(Keywords_p), c=0: parsing(a, b, c)
        parsed_content = x()
        prof=""
        try:
            for k, v in parsed_content.items():
                if (k == "Professional Summary") or (k == "Experience Summary") or (k == "EXPERIENCE SUMMARY") or \
                        (k == "PROFESSIONAL SUMMARY") or (k == "SUMMARY") or (k == "OBJECTIVE"):
                    prof = v.replace(k,"")
                    prof=re.sub("[:]","",prof)
                    prof=" ".join(prof.split())
        except:
            prof = "Not mentioned"
        prof = prof
        prof_summary = re.sub("Professional Summary+""[:,-]", "", prof)

        # Total years of experience
        profess=prof
        experience = None
        try:
            if profess != "Not mentioned":
                if "experience" in profess.lower():
                    exp_index=profess.index("experience")
                    profess_val=(prof[exp_index - 20:exp_index+40])
                    exp_list = re.findall('[0-9]+', profess_val[0:20])
                    exp = ' '.join(str(i) for i in exp_list)
                    experience = exp
        except:
            experience = None

        # To extract name
        test = spacy.blank('en')
        ts = test(" ".join(text.split('\n')))
        name=""
        for i in ts[0:4]:
            name+=str(i)
        name=re.split(" ",name)
        if "." in name:
            name.remove(".")

        res = next(sub for sub in name if sub)      # To print first not null value
        Name=res
        if Name in Keywords:
            Name = None

        return {"Name":Name,"Email": email, "Phone_number": phone_number, "Year Of Experience": experience,
                "Designation": designation, "Location": loc, "Skills": skill, "Secondary skills": sskill,
                "Passed_out":passed_out,"Qualifcation":qualifcation,
                "Bachelor Education": bachelor_degree, "Bachelor College": bachelor_college,
                "Master Education": mast_deg,
                "Master College": mast_college,"Project":project, "Education History": education,"professional_summary":prof_summary
                }
    else:
        resp = jsonify(errors)
        resp.status_code = 500
        return resp


if __name__ == "__main__":
    app.run(debug=True)