from __future__ import print_function
import time
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException
from pprint import pprint


def convert_to_pdf(path_to_odt_document):
    configuration = cloudmersive_convert_api_client.Configuration()
    configuration.api_key['Apikey'] = 'f0c513bc-8c00-4491-830e-3e83b015feb6'

    api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))

    try:
        api_response = api_instance.convert_document_odt_to_pdf(path_to_odt_document)
        pprint(api_response)
        return api_response
    except ApiException as e:
        print("Exception when calling ConvertDocumentApi->convert_document_odt_to_pdf: %s\n" % e)

