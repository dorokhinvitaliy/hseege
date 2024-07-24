# This is a sample Python script.
from flask_cors import cross_origin
from flask import Flask, jsonify, request
from createExcel import createExcel
from parseExcel import parseExcel
from parseHSE import parseSelenium
from parseSnils import parseSnils

from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from models import db, InfoModel

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


vuzes = [
    {"id": 1, "name": "ВШЭ", "naprs": [
        {"id": 1, "napr": "Бизнес-информатика"},
        {"id": 2, "napr": "Прикладная математика и информатика"},
        {"id": 3, "napr": "Программная инженерия"},
        {"id": 4, "napr": "Информатика и вычислительная техника"},
        {"id": 5, "napr": "Информационная безопасность"},
    ]}
]
naprs = [i["napr"] for i in vuzes[0]["naprs"]]
# Press the green button in the gutter to run the script.


app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False
app.config['CORS_HEADERS'] = 'Content-Type'

app.config['SQLALCHEMY_DATABASE_URI'] = "postgresql://<postgres>:<>@<localhost>:5432/<cotest>"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db.init_app(app)
migrate = Migrate(app, db)


@app.route('/login', methods=['POST', 'GET'])
def login():
    if request.method == 'GET':
        return "Login via the login Form"

    if request.method == 'POST':
        name = request.form['name']
        age = request.form['age']
        new_user = InfoModel(name=name, age=age)
        db.session.add(new_user)
        db.session.commit()
        return f"Done!!"

@app.route('/ege', methods=["POST"])
@cross_origin(supports_credentials=True)
def get_ege():
    napr = request.json.get('napr')
    vuz = request.json.get('vuz')
    prior = request.json.get('prior')
    att = request.json.get('attestat')
    table = parseExcel(vuzes[int(vuz) - 1]["name"],
                       vuzes[int(vuz) - 1]["naprs"][int(napr) - 1]["napr"],
                       int(prior),
                       int(att))
    return jsonify({'reit': table[0], 'b': table[1], 'p': table[2], "name": table[3], "version_date": table[4]})


@app.route('/ege/file', methods=['GET'])

def get_file():
    table = parseExcel()
    createExcel(table[3], table[0][0], table[0][1], table[0][2], table[1], table[2])
    return jsonify({"name": "vanya"})


@app.route('/ege/parse', methods=['GET'])
@cross_origin(supports_credentials=True)
def get_parsing():
    parseSelenium(naprs)
    return jsonify({"success": True})


@app.route('/ege/snils', methods=['POST'])
@cross_origin(supports_credentials=True)
def get_snils():
    res = parseSnils("ВШЭ", naprs, 25, 0, request.json.get("snils"))
    return jsonify(res)


if __name__ == '__main__':
    app.run(debug=True)
    # parseExcel()
