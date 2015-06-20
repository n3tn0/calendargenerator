main_auto.py : sumy.ui
	pyuic5 sumy.ui -o main_auto.py

dist/Sumy-0.1-amd64.msi : main.py
    python cxfreeze main.py
