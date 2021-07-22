# SHG_analysis
Scripts that speed up data analysis for SHG imaging of E5 chick hearts

Script tips:
Download anaconda(https://docs.anaconda.com/anaconda/install/index.html) and run using spyder, the script was written in spyder in a conda enviroment 
so standalone spdyer programs may not work


For reading data:
You can eliminate the lines related to user input (at the top). Eliminate the working directory line (it'll save the made files in the the same folder as wherever the python script is stored). You can hard code file path in the line that asks for a file path and make sure the eliminate the .format() at the end of strings (basically words with quotations around them) in lines for the file path. There is also a user input that asks for z slices, so just
be careful if you hard code other parts of the user input. 

