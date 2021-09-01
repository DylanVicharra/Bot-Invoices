
class Orden:
    def __init__(self, email, password, orden):
        self.email = email
        self.password = password
        self.orden = orden
        self.serials = []
        self.estado = None