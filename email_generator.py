from jinja2 import Environment, PackageLoader, select_autoescape
from jinja2.runtime import to_string
import pandas as pd
from pandas.core.reshape.concat import concat
from datetime import datetime

class TemplateProcessor:
    def __init__(self,template_file, **kwargs):
        """Initiates a template using a jinja template file."""
 
        env = Environment(
            loader=PackageLoader('email_generator', 'templates'),
            autoescape=select_autoescape(['html', 'xml'])
        )
        template = env.get_template(template_file)
        self.body = template.render(**kwargs)
        

    def save_file(self, filename):
        """Saves the body of the template object in a file with the filename as parameter"""
       
        f = open(filename,'w')
        f.write(self.body)
        f.close 
    
    def send_email(self,to, subj):
        """Sends an email using the template object."""

        """ Currently not implemented. """
        print(to+'-'+subj)
        print(self.body)
        # Send the finalized email here.

def main():

    data = pd.read_excel('receiver data.xlsx', sheet_name= 'Sheet1', )
   
    data['owners_city_state_zipcode'] = data['OWNER_CITY'] + ','+ data['OWNER_STATE'] + ','+ data['OWNER_ZIP'].astype(int).astype(str)
    data['site_address_city_zip'] = data['SITE_ADDR'] + ','+ data['SITE_CITY'] + ','+ data['SITE_ZIP'].astype(int).astype(str)
    data['owners_name'] = data['OWNER_1_FIRST'] + ' '+ data['OWNER_1_LAST']
    
    
    for index, row in data.iterrows():
        template = TemplateProcessor(template_file='Letter_Template.html',
        owners_address=row['OWNER_ADDRESS'],
        owners_city_state_zipcode=row['owners_city_state_zipcode'],
        date=datetime.today().strftime('%d/%m/%Y'),
        owners_name=row['owners_name'],
        site_address_city_zip=row['site_address_city_zip']
        )
        filename = str(index) + row['owners_name'] + '.html'
        template.save_file(filename)

   
    # template.send_email(
    #     to='test@email.com',
    #     subj='test subject',
    #     )

if __name__ == main():
    main()