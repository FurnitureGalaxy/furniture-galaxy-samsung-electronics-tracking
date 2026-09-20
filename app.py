import os, json
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash

app=Flask(__name__)
app.secret_key=os.environ.get("SECRET_KEY","fg-samsung-order-tracking-change-me")
DATA_DIR=Path(os.environ.get("DATA_DIR","data"))
DATA_DIR.mkdir(parents=True,exist_ok=True)
DATA_FILE=DATA_DIR/"current_orders.xlsx"
META_FILE=DATA_DIR/"metadata.json"

# Samsung file uses fixed Excel columns:
# E = PO NUMBER, K = QUANTITY ORDERED, M = MODEL NUMBER, S = ESTIMATED SHIP DATE
COLUMN_INDEXES={"PO Number":4,"Quantity Ordered":10,"Model Number":12,"Estimated Ship Date":18}

def read_samsung_excel(path):
    # Read raw sheet so Samsung can change header wording without breaking the page.
    raw=pd.read_excel(path,dtype=object)
    if raw.shape[1] < 19:
        raise ValueError("This file does not contain Column S. Please upload the Samsung orderTracking .xlsx file.")
    out=pd.DataFrame()
    for name,idx in COLUMN_INDEXES.items():
        out[name]=raw.iloc[:,idx]
    for c in ["PO Number","Quantity Ordered","Model Number"]:
        out[c]=out[c].fillna("").astype(str).str.strip().replace({"nan":"","None":""})
    # Keep date as supplied, but present dates cleanly. DO NOT add 10 days.
    def fmt_date(v):
        if pd.isna(v): return ""
        if isinstance(v,(pd.Timestamp,datetime)):
            return v.strftime("%Y-%m-%d")
        return str(v).strip()
    out["Estimated Ship Date"]=out["Estimated Ship Date"].apply(fmt_date)
    return out[(out["PO Number"]!="") | (out["Model Number"]!="")].reset_index(drop=True)

def load_data():
    if not DATA_FILE.exists():
        return pd.DataFrame(columns=list(COLUMN_INDEXES))
    return read_samsung_excel(DATA_FILE)

def last_updated():
    if META_FILE.exists():
        try:
            return json.loads(META_FILE.read_text(encoding="utf-8")).get("last_updated","")
        except Exception:
            pass
    return ""

@app.route("/",methods=["GET"])
def index():
    po=request.args.get("po","").strip()
    model=request.args.get("model","").strip()
    try:
        df=load_data()
        if po:
            df=df[df["PO Number"].str.contains(po,case=False,na=False,regex=False)]
        if model:
            df=df[df["Model Number"].str.contains(model,case=False,na=False,regex=False)]
        rows=df.to_dict("records")
        error=None
    except Exception as e:
        rows=[]; error=str(e)
    return render_template("index.html",rows=rows,po=po,model=model,last_updated=last_updated(),error=error)

@app.route("/upload",methods=["POST"])
def upload():
    f=request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Please choose a Samsung .xlsx orderTracking file.")
        return redirect(url_for("index"))
    temp=DATA_DIR/"upload_check.xlsx"
    f.save(temp)
    try:
        test=read_samsung_excel(temp)
        if len(test)==0:
            raise ValueError("No Samsung order rows were found in Columns E, K, M and S.")
        temp.replace(DATA_FILE)
        stamp=datetime.now(ZoneInfo("America/Edmonton")).strftime("%B %d, %Y at %I:%M %p")
        META_FILE.write_text(json.dumps({"last_updated":stamp}),encoding="utf-8")
        flash(f"Samsung Order Tracking updated successfully — {len(test):,} row(s) loaded.")
    except Exception as e:
        if temp.exists(): temp.unlink()
        flash("Upload failed: "+str(e))
    return redirect(url_for("index"))

@app.route("/health")
def health():
    return {"status":"ok"},200

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT","10000")))
