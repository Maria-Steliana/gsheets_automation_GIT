import datetime
import json

from gsheets import GoogleSheets
from mail import Mail

# o sa avem gsheets cu nume, mail, data expirare parola
# vom face automatizare care trimite mailuri angajatilor carora urmeaza sa le expire parola


def read_config(path: str = "config.json") -> dict:
    with open(path, "r") as f:
        config = json.loads(f.read())
        return config


def read_html_template(path: str = "mail_template.html") -> str:
    with open(path, "r") as f:
        html_template = f.read()
        return html_template


def run(employees: list, mail: Mail, config: dict):
    today = datetime.datetime.today()

    for employee in employees:
        expiration_date = datetime.datetime.strptime(employee['Password Expiration Date'], "%d/%m/%Y")
        delta_days = (expiration_date - today).days
        if 3 >= delta_days >= 0:
            print(f"User {employee['Name']} needs to reset his password. Mail will be sent")
            html_template = read_html_template()
            html_template = html_template.replace("[Recipient's Name]", employee['Name'])
            if delta_days < 1:
                html_template = html_template.replace("[color_background]", config["color_codes"]["red"])
            elif delta_days < 2:
                html_template = html_template.replace("[color_background]", config["color_codes"]["yellow"])
            else:
                html_template = html_template.replace("[color_background]", config["color_codes"]["green"])

            mail.send_email_using_mime(employee['Mail'], "Password about to expire", html_template)

        else:
            print(f"User {employee['Name']} doesn't need to reset his password.")


if __name__ == '__main__':
    config = read_config()
    mail = Mail(config['sender_email'])
    excel = GoogleSheets(config['scopes'], "credentials.json", config['excel_id'])
    employees = excel.get_values()
    run(employees, mail, config)


