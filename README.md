# excelsearchdocx
1. useage
read keyword from excel,search from doc docx dir.

2. 用 Pyinstaller 打包 Python exe程序
#use pipenv
pip install pipenv
pipenv install --python 3.7
pipenv shell
pip list


#install
pipenv install xlwt
pipenv install xlrd
pipenv install lxml
pipenv install python-docx
pipenv install pypiwin32

#pyinstaller,use your ico
pyinstaller --distpath Release/ -w -i favicon.ico --clean tk2_searchdocx_by_excelkeyword.py
