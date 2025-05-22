from flask import Flask, request, jsonify
import os
import csv

app = Flask(__name__)

# CSV storage file
CSV_FILE = 'notes.csv'

def get_notes():
    notes = []
    if os.path.exists(CSV_FILE):
        with open(CSV_FILE, 'r', newline='') as f:
            reader = csv.reader(f)
            notes = [row[0] for row in reader if row]  # Get first column of each row
    return notes

def save_notes(notes):
    with open(CSV_FILE, 'w', newline='') as f:
        writer = csv.writer(f)
        for note in notes:
            writer.writerow([note])  # Write each note as a single-column row

@app.route('/')
def home():
    return open('index.html').read()

@app.route('/save', methods=['POST'])
def save():
    note = request.json.get('note', '')
    if note:
        notes = get_notes()
        notes.append(note)
        save_notes(notes)
        return jsonify({'success': True})
    return jsonify({'success': False})

@app.route('/load', methods=['GET'])
def load():
    return jsonify({'notes': get_notes()})

@app.route('/delete', methods=['POST'])
def delete():
    index = request.json.get('index', -1)
    notes = get_notes()
    if 0 <= index < len(notes):
        notes.pop(index)
        save_notes(notes)
        return jsonify({'success': True})
    return jsonify({'success': False})

if __name__ == '__main__':
    app.run(debug=True, port=5000)