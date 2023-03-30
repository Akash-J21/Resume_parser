In this python program, Resume parsing will be done using the python version 3.0

The code is developed in pycharm community environment.
The API used here is POSTMAN to GET and POST  
Note: .pdf and .docx are the acceptable resume format

Using this resume parser user can get Name, Email, Phone_number,Year Of Experience, Designation, Location, Skills, Secondary skills, Passed_out, Qualifcation, Bachelor Education, Bachelor College, Master Education, Master College, Project, Education History and professional_summary
                

Step 1: Once the code is copied, set your local path in the UPLOAD_FOLDER variable where the 	document is stored when send it through API POST

Step 2: Copy the same path in res_model variable to get back the sam file from locally stored 	memory 

Step 3: Run the code with GET method to test the API connection.

Step 4: After verifying the connection, using POST method open the API and choose form-data 	followed by entering Key and Value data.
	set Key as files[] 
	set Value as the pdf or docx file which user need to parse
Note: API used here is POSTMAN

Step 5: Send the POST to get result in JSON format .