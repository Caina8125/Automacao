class UserLogin:
    _user: str
    _password: str

    def __init__(self, user: str, password: str) -> None:
        self._user = user
        self._password = password
    
    @property
    def user(self):
        return self._user
    
    @user.setter
    def user(self, user: str):
        self.user = user

    @property
    def password(self):
        return self._password
    
    @password.setter
    def password(self, password: str) -> None:
        self.password = password