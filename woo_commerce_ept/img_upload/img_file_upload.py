# -*- coding: utf-8 -*-
# See LICENSE file for full copyright and licensing details.
'''
In this file transfer images from odoo to woo-commerce.
'''
import tempfile
import base64
import requests
from odoo.tools.mimetypes import guess_mimetype
# from .. python_magic_0_4_11 import magic
from ..wordpress_xmlrpc import base
from ..wordpress_xmlrpc import compat
from ..wordpress_xmlrpc import media


class SpecialTransport(compat.xmlrpc_client.Transport):
    '''
    In this class transfer images from odoo to woo-commerce.
    '''
    user_agent = 'Mozilla/5.0 (Windows NT 6.0) AppleWebKit/537.31 (KHTML, like Gecko) Chrome/26.0.1410.43 Safari/537.31'


def upload_image(instance, image_data, image_name, mime_type=False):
    '''

    :param instance:  Woo-active instance
    :param image_data: Image Binary Data
    :param image_name: Image Name
    :param mime_type: Image Mime Type If Image have set.
    :return: Its pass request of image in woo-commerce
    '''
    if instance.woo_admin_username and instance.woo_admin_password:

        if not image_data or not image_name:
            return {}

        client = base.Client('%s/xmlrpc.php' % (instance.woo_host), instance.woo_admin_username,
                             instance.woo_admin_password, transport=SpecialTransport())

        binary_data = base64.decodebytes(image_data)

        data = {
            'name': '%s_%s.%s' % (image_name, instance.id, mime_type.split("/")[1]),
            'type': mime_type,
            'bits': compat.xmlrpc_client.Binary(binary_data)
        }
        res = client.call(media.UploadFile(data))

        return res
    return {}


def fetch_image(image_url):
    '''

    :param image_url: ImageURL
    :return: If get image then image content or False
    '''
    if not image_url:
        return False
    try:
        img = requests.get(image_url, stream=True, timeout=10)
    except:
        img = False
    return img and base64.b64encode(img.content) or False
