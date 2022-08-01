import random
import statistics
from typing import Any, Dict

from faker.providers import BaseProvider
from .measures import measures

class BloodPanelProvider(BaseProvider):
    '''
    A Provider that generates mock blood panel values.

    >>> from faker import Faker
    >>> from faker_bloodpanel import BloodPanelProvider
    >>> fake = Faker()
    >>> fake.add_provider(BloodPanelProvider)
    '''

    def simple_blood_panel(self, use_long_names: bool = True, use_short_names: bool = False, add_units: bool = False) -> Dict[str, Any]:
        '''
        Returns a simple blood panel with a few values as a dictionary:
        '''
        simple_measures = ['sodium', 'potassium', 'chloride', 'bicarbonate', 'urea', 'magnesium', 'calcium', 'hemoglobin']
        bp = {}
        for i, simple_measure in enumerate(simple_measures):
            meas_info = measures[simple_measure]

            name = f'{i}'
            if use_long_names:
                name = meas_info['long_name']
            if use_short_names and ('short_name' in meas_info.keys()):
                if use_long_names:
                    name = f'{name} ({meas_info["short_name"]})'
                else:
                    name = meas_info['short_name']

            mu = 1.0 * (meas_info['lower'] + meas_info['upper']) / 2.0
            sigma = (mu - meas_info['lower']) / 2
            value = f'{random.gauss(mu, sigma):.2f}'
            if add_units:
                value = f'{value} {meas_info["unit"]}'

            bp[name] = value
        return bp

    def covid_test(self) -> Dict[str, Any]:
        '''
        Returns a simple covid test answer as a simple dictionary containing a boolean value.
        '''
        return { 'Covid-PCR': True if random.uniform(0.0, 1.0) < .2 else False }
