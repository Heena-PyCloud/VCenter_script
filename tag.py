import ssl
import time
import requests
import pandas as pd
from com.vmware.cis.tagging_client import (
    Category, Tag, TagAssociation, CategoryModel)
from com.vmware.cis_client import Session
from com.vmware.vapi.std_client import DynamicID
from pyVim import connect
from pyVmomi import vim
from vmware.vapi.lib.connect import get_requests_connector
from vmware.vapi.security.session import create_session_security_context
from vmware.vapi.security.user_password import \
    create_user_password_security_context
from vmware.vapi.stdlib.client.factories import StubConfigurationFactory
requests.packages.urllib3.disable_warnings()
server = '10.#.#.#'
url = 'https://{}/api'.format(server)
username = '****'
password = '*****'
path = '/home/user1/Downloads/Test.xlsx'
"""
Create an authenticated stub configuration object that can be used to issue
requests against vCenter.
Returns a stub_config that stores the session identifier that can be used
to issue authenticated requests against vCenter.
"""
session = requests.Session()
session.verify = False
connector = get_requests_connector(session=session, url=url)
stub_config = StubConfigurationFactory.new_std_configuration(connector)

# Pass user credentials (user/password) in the security context to authenticate.
# login to vAPI endpoint
user_password_security_context = create_user_password_security_context(username,
                                                                       password)
stub_config.connector.set_security_context(user_password_security_context)

# Create the stub for the session service and login by creating a session.
session_svc = Session(stub_config)
session_id = session_svc.create()

# Successful authentication.  Store the session identifier in the security
# context of the stub and use that for all subsequent remote requests
session_security_context = create_session_security_context(session_id)
stub_config.connector.set_security_context(session_security_context)

# Create Tagging services
tag_svc = Tag(stub_config)
category_svc = Category(stub_config)
tag_association = TagAssociation(stub_config)

def get_vm_id(name):
    """Find vm id by given name using pyVmomi."""
    context = None
    if hasattr(ssl, '_create_unverified_context'):
        context = ssl._create_unverified_context()

    si = connect.Connect(host=server, user=username, pwd=password,
                         sslContext=context)
    content = si.content
    container = content.rootFolder
    viewType = [vim.VirtualMachine]
    recursive = True
    vmView = content.viewManager.CreateContainerView(container,
                                                     viewType,
                                                     recursive)
    vms = vmView.view
    for vm in vms:
        if vm.name == name:
           return vm._GetMoId()
    raise Exception('VM with name {} could not be found'.format(name))


def tag_id(tag_name):
   tags = tag_svc.list()

   for tag in tags:
      t = tag_svc.get(tag)
      if t.name.lower() == tag_name.lower():
         #print(t.id)
         return {
                "id": t.id,
                "category_id": t.category_id
            }

def append_tag():

   """
   Read the excel file in pandas dataframe and output name and tag
   get the vm-id and tag-id
   attach the tag to vm
   """
   df = pd.read_excel(path)
   df1=df[['VM Name','Taged to']]  #Reading the VM Name and tag column from the sheet
   df2=df1.rename(columns={'VM Name': 'name','Taged to': 'tag_name'})
   for ind in df2.index:
      name=df2['name'][ind]
      tag_name_col=df2['tag_name'][ind]
      tag_name_split=tag_name_col.split(',')
      vm_moreof=get_vm_id(name)
      for tag_name in tag_name_split:
         tag=tag_id(tag_name)
         if tag is not None:			
            dynamic_id = DynamicID(type='VirtualMachine', id=vm_moreof)
            tag_association.attach(
            tag_id=tag['id'], object_id=dynamic_id)
            for tag['id'] in tag_association.list_attached_tags(dynamic_id):
               if tag['id'] == tag['id']:
                  tag_attached = True
                  break
            assert tag_attached
            print('Tagged vm: {0}'.format(vm_moreof))

append_tag()

