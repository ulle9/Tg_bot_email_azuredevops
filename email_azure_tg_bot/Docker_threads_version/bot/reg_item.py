from azure.devops.connection import Connection
from msrest.authentication import BasicAuthentication
from azure.devops.v7_0.work_item_tracking.models import JsonPatchOperation
import configparser


# PAT, org URL and templates dictionary
config = configparser.ConfigParser()
config.read_file(open(r'e_coll.cfg', encoding='utf-8'))
personal_access_token = config.get('reg_item', 'personal_access_token')
organization_url = config.get('reg_item', 'organization_url')
template_dict = dict()
for i in config.get('reg_item', 'template_dict').split(';'):
    template_dict[i.split(':')[0]] = (i.split(':')[1].split(',')[0], int(i.split(':')[1].split(',')[1]))


# Register problem in Azure DevOps
def register_problem(project, time, title, text):
    project_name = template_dict[project][0]
    template_num = template_dict[project][1]

    # Create a connection to the org
    credentials = BasicAuthentication('', personal_access_token)
    connection = Connection(base_url=organization_url, creds=credentials)

    # Get template work item for the project
    wit_client = connection.clients.get_work_item_tracking_client()
    template_item = wit_client.get_work_item(template_num)

    # Create fields for new work item
    title = JsonPatchOperation(op='add', path='/fields/System.Title', value='[Text!]: {}, {}'.format(title, time))
    support_type = JsonPatchOperation(op='add', path='/fields/Text.SupportType', value=template_item.fields['Text.SupportType'])
    description = JsonPatchOperation(op='add', path='/fields/System.Description', value=text+template_item.fields['System.Description'])
    assign_to = JsonPatchOperation(op='add', path='/fields/System.AssignedTo', value=template_item.fields['System.AssignedTo'])
    area_path = JsonPatchOperation(op='add', path='/fields/System.AreaPath', value=template_item.fields['System.AreaPath'])
    iter_path = JsonPatchOperation(op='add', path='/fields/System.IterationPath', value=template_item.fields['System.IterationPath'])
    if 'Text.Problem.Chronology' in template_item.fields:
        chronology = JsonPatchOperation(op='add', path='/fields/Text.Problem.Chronology', value=template_item.fields['Text.Problem.Chronology'])
    else:
        chronology = JsonPatchOperation(op='add', path='/fields/Text.Problem.Chronology', value='Chronology')

    # Create work item (problem)
    create_wit = wit_client.create_work_item([title, support_type, description, chronology, assign_to, area_path, iter_path], project_name, 'Problem')

    return create_wit.id

