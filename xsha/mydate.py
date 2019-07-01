import datetime


class Holiday:
    def __init__(self,date=datetime.datetime.today()):
        self.date=date
        self.isweekend=self._check_week()
        # self.holiday=self._get_holiday()
        # self.workday=self._get_workday()

    def _check_week(self):
        if self.date.weekday() in [5,6]:
            return True
        else:
            return False

    def __bool__(self):
        return self.isweekend

    def __str__(self):
        return "{}年{}月{}日，{}假期！".format(
            self.date.year,
            self.date.month,
            self.date.day,
            "是" if self.isweekend else "不是")

    def __repr__(self):
        return "Holiday(date=datetime.date({},{},{}))".format(
            self.date.year,
            self.date.month,
            self.date.day)


if __name__ == '__main__':
    if Holiday(datetime.date(2019, 6, 7)):
        print('True')
    else:
        print('False')

    print(Holiday())
