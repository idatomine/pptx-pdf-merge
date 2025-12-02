# PPTX to Merged PDF Converter (Flask + LibreOffice)

Er du også lei av at det alltid skal koste penger å merge sammen hundrevis av powerpointer når du tar skippertak tre dager før eksamen? Prøv denne knalljbra løsningen som er et resultat av at jeg prokrastinerer tre dager før eksamen. 
  
Laget med **Flask**, **LibreOffice (soffice)**, **PyPDF2** og kjærlighet 🥰


## 🖥 Requirements

Du trenger:

- **Python 3.9+**
- **LibreOffice** innstallert
- Pip pakkene fra `requirements.txt`

---

## 📦 Installation

Below is the complete setup guide for **macOS**

### 1 Install Python 3 
Hvis du ikke har python på pcen din så burde du lære deg det først

### 2 Install LibreOffice (required for PPTX→PDF conversion)
```
brew install --cask libreoffice
```

Default macOS path:
```
/Applications/LibreOffice.app/Contents/MacOS/soffice
```

### 3 Install Python dependencies
```
pip install -r requirements.txt
```

### 4 Set the correct soffice path in app.py
```
SOFFICE_CMD = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
```


▶️ Running the app
Start the development server:
```
python app.py
```

You should see:
 * Running on http://127.0.0.1:5000

Open your browser at:

http://127.0.0.1:5000

# NB!!! Bruk med omhu, dette kan gjøre deg til et akademisk våpen på rekordtid🔥