from os import path
from datetime import datetime


class Logger:
    def __init__(self, filename):
        self.errors = []
        self.filename = path.join(filename)

    def append_error(self, error):
        self.errors.append(error)

    def count_errors(self):
        return len(self.errors)

    def prepare_output(self):
        output = "\n==================\n"
        errors = "\n- ".join(self.errors)
        date = datetime.today()
        cnt_errors = self.count_errors()
        if errors:
            if cnt_errors == 1:
                output += "Завершено {0} с {1} ошибкой: \n- {2}".format(date, cnt_errors, errors)
            else:
                output += "Завершено {0} с {1} ошибками: \n- {2}".format(date, cnt_errors, errors)
        else:
            output += "Завершено {0} без ошибок.".format(date)
        return output

    def save(self):
        output = self.prepare_output()
        with open(self.filename, "a", encoding='utf8') as file:
            file.write(output)
