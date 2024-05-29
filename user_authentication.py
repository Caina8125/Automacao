class UserLogin:
    def __init__(self, user: str, password: str) -> None:
        if user == 'pedro.pereira' or 'PEDRO.PEREIRA':
            self._user = 'faturamento.fat'
            self._password = '87316812#hg12@'
        else:    
            self._user = user
            self._password = password
    
    @property
    def user(self):
        return self._user
    
    @user.setter
    def user(self, user: str):
        self._user = user

    @property
    def password(self):
        return self._password
    
    @password.setter
    def password(self, password: str) -> None:
        self._password = password