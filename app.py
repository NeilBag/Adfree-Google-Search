from flask import Flask, render_template, request, send_file, jsonify
import requests
from bs4 import BeautifulSoup
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import io
import openpyxl
import xlsxwriter

app = Flask(__name__)

email_config = {
    'host': 'smtp.gmail.com',
    'port': 587,
    'secure': False,
    'auth': {
        'user': 'bagchineil518@gmail.com',
        'pass': 'vyyq qbgl ckat gttp',
    },
}

def fetch_search_results(query):
    google_search_url = f"https://www.google.com/search?q={query}&num=500"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    response = requests.get(google_search_url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')

    search_results = []
    for result in soup.select('.tF2Cxc'):
        title_tag = result.select_one('h3')
        link_tag = result.select_one('.yuRUbf a')
        snippet_tag = result.select_one('.IsZvec')

        title = title_tag.text if title_tag else 'No title available'
        link = link_tag['href'] if link_tag else 'No link available'
        snippet = snippet_tag.text if snippet_tag else 'No snippet available'

        search_results.append({
            'Title': title,
            'Link': link,
            'Snippet': snippet
        })
    
    return search_results

def save_results_to_excel(results):
    df = pd.DataFrame(results)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Results', index=False)
    writer.save()
    output.seek(0)
    return output

def send_email(results, recipient_email, search_query):
    msg = MIMEMultipart()
    msg['From'] = email_config['auth']['user']
    msg['To'] = recipient_email
    msg['Subject'] = f"Google Search Results for: {search_query}"

    html_table = """
    <table border="1">
      <thead>
        <tr>
          <th>Title</th>
          <th>Link</th>
          <th>Snippet</th>
        </tr>
      </thead>
      <tbody>
    """
    for result in results:
        html_table += f"""
        <tr>
            <td>{result['Title']}</td>
            <td><a href="{result['Link']}" target="_blank">{result['Link']}</a></td>
            <td>{result['Snippet']}</td>
        </tr>
        """
    html_table += """
      </tbody>
    </table>
    """
    msg.attach(MIMEText(html_table, 'html'))

    excel_buffer = save_results_to_excel(results)
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(excel_buffer.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename=GoogleSearchResults.xlsx')
    msg.attach(part)

    with smtplib.SMTP(email_config['host'], email_config['port']) as server:
        server.starttls()
        server.login(email_config['auth']['user'], email_config['auth']['pass'])
        server.sendmail(msg['From'], msg['To'], msg.as_string())

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/search', methods=['POST'])
def search_route():
    try:
        search_query = request.json['searchQuery']
        results = fetch_search_results(search_query)
        return jsonify(results)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/excel', methods=['POST'])
def download_excel():
    try:
        results = request.json['results']
        excel_buffer = save_results_to_excel(results)
        return send_file(
            excel_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            attachment_filename='GoogleSearchResults.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/email', methods=['POST'])
def email_results():
    try:
        results = request.json['results']
        recipient_email = request.json['recipientEmail']
        search_query = request.json['searchQuery']
        send_email(results, recipient_email, search_query)
        return jsonify({'message': 'Email sent successfully.'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
