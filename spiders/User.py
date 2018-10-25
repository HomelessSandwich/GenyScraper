from datetime import datetime
from getpass import getpass
import unittest


class User():
    """
    This class describes a user for the website www.pmu.fr
    """

    def __repr__(self):
        return f'Customer Number: {self.customer_num}, PIN: {self.pin}, Birthday: {self.birthday}'

    @property
    def customer_num(self):
        """User's customer number."""
        return self._customer_num

    @customer_num.setter
    def customer_num(self, num):
        if len(num) == 10:
            self._customer_num = num
        else:
            raise ValueError('Not a valid customer number!')

    @property
    def pin(self):
        """User's PIN number."""
        return self._pin

    @pin.setter
    def pin(self, pin):
        if len(pin) == 6:
            self._pin = pin
        else:
            raise ValueError('Not a valid PIN number!')

    @property
    def birthday(self):
        """User's brithday. Needs to be in the DD/MM/YYYY format."""
        return self._birthday

    @birthday.setter
    def birthday(self, birthday):
        if self.validate_date(birthday):
            self._birthday = birthday
        else:
            raise ValueError('Not a valid date!')

    @property
    def birth_date(self):
        """User's day of birth."""
        return self.birthday.split('/')[0]

    @property
    def birth_month(self):
        """User's month of birth."""
        return self.birthday.split('/')[1]

    @property
    def birth_year(self):
        """User's year of birth."""
        return self.birthday.split('/')[2]

    @staticmethod
    def validate_date(date):
        """Validates a given date. Needs to be in the DD/MM/YYYY format to be excepted."""
        try:
            datetime.strptime(date, '%d/%m/%Y')
        except ValueError:
            return False
        else:
            return True

    def get_user_details(self):
        """Asks the user for their details. Deals with error handling."""
        # Birthday
        while True:
            try:
                self.birthday = input("Birthday (DD/MM/YYYY): ")
            except ValueError:
                print('\nThat was not was valid birthday!')
            else:
                break


class TestUserMethods(unittest.TestCase):
    """Unit testing for the User class."""

    def setup(self):
        user = User()


if __name__ == '__main__':
    unittest.main()
