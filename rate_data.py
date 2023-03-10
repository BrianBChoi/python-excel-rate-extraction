class RateTable:
    def __init__(self, name):
        self.name = name
        self.commodity_codes = []
        self.origins = []
    
    def add_commodity_code(self, code):
        self.commodity_codes.append(code)

    def add_origin(self, origin):
        self.origins.append(origin)


class Origin:
    def __init__(self, name):
        self.name = name
        self.destinations = []
    
    def add_destination(self, destination):
        self.destinations.append(destination)


class Destination:
    def __init__(self, name, term, via):
        self.name = name
        self.term = term
        self.via = via
        self.rates = {}

    def add_rate(self, container_type, rate):
        self.rates[container_type] = rate