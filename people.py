class Person ():

    def __init__(self, name):
        self.name = name
        self.dialect = None

    def __eq__(self, query):
        return query.lower() in self.name.lower()

    def __repr__(self):
        return self.get_name()

    def set_dialect(self, lang):
        self.dialect = lang

    def get_name(self):
        return self.name


class Elder (Person):

    def __init__(self, name, blk, floor, unit, hosp):
        Person.__init__(self, name)
        self.nickname = None
        self.blk, self.floor, self.unit = blk, floor, unit
        self.note = ''
        self.hospitability = hosp
        self.suspended = False

    def set_nickname(self, nickname):
        self.nickname = nickname

    def add_comment(self, txt):
        self.note = txt

    def get_hosp(self):
        return self.hospitability

    def suspend(self):
        self.suspended = True

    def unsuspend(self):
        self.suspended = False


class Volunteer (Person):

    def __init__(self, name):
        Person.__init__(self, name)
        self.availability = []  # [dates]
        self.attendance = []  # [dates]
        self.pulled_out = []  # [dates]
        self.visited = {}  # name: int
        self.friend = {}  # name :int

    def add_avail(self, date):
        self.availability.append(date)

    def is_available(self, date):
        return date in self.availability

    def attended(self, date):
        self.attendance.append(date)

    def pull_out(self, date):
        self.pulled_out.append(date)

    def pull_out_rate(self):
        try:
            return len(self.pulled_out) / (len(self.attendance) + len(self.pulled_out))
        except ZeroDivisionError:
            return 0

    def visits(self, elder):
        if elder not in self.visited:
            self.visited[elder] = 1
        else:
            self.visited[elder] += 1

    def add_friend(self, name):
        if name not in self.friend:
            self.friend[name] = 1
        else:
            self.friend[name] += 1

    def times_visited(self, query):
        if query in self.visited:
            return self.visited[query]
        else:
            return 0
