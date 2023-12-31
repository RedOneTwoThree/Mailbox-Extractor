import json
from configparser import SectionProxy
from arrow import get
from azure.identity import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.user_item_request_builder import (
    UserItemRequestBuilder,
)
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from msgraph.generated.users.item.mail_folders.mail_folders_request_builder import (
    MailFoldersRequestBuilder,
)

from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody,
)
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress


class Graph:
    settings: SectionProxy
    device_code_credential: DeviceCodeCredential
    user_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings["clientId"]
        tenant_id = self.settings["tenantId"]
        graph_scopes = self.settings["graphUserScopes"].split(" ")

        self.device_code_credential = DeviceCodeCredential(
            client_id, tenant_id=tenant_id
        )
        self.user_client = GraphServiceClient(self.device_code_credential, graph_scopes)

    async def get_user_token(self):
        graph_scopes = self.settings["graphUserScopes"]
        access_token = self.device_code_credential.get_token(graph_scopes)
        return access_token.token

    async def get_user(self):
        # Only request specific properties using $select
        query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
            select=["displayName", "mail", "userPrincipalName"]
        )
        request_config = (
            UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
        )
        user = await self.user_client.me.get(request_configuration=request_config)
        return user

    async def get_inbox(self):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            # Only request specific properties
            select=["from", "isRead", "receivedDateTime", "subject"],
            # Get at most 25 results
            top=25,
            # Sort by received time, newest first
            orderby=["receivedDateTime DESC"],
        )
        request_config = (
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
        )

        messages = await self.user_client.me.mail_folders.by_mail_folder_id(
            "inbox"
        ).messages.get(request_configuration=request_config)
        return messages

    async def get_folder(self):
        query_params = (
            MailFoldersRequestBuilder.MailFoldersRequestBuilderGetQueryParameters(
                top=100
            )
        )

        request_config = (
            MailFoldersRequestBuilder.MailFoldersRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
        )

        folders = await self.user_client.me.mail_folders.by_mail_folder_id(
            "AAMkADZhZmQwMDBiLTJjNTEtNDg4Zi1iYjQ1LTdhYmQ4NTFjZmZiMQAuAAAAAAB9ICr57bFCQYIG8XhpgM7FAQDoVoEaFdnbQpsKukhxf4gBAAAB1hKlAAA="
        ).child_folders.get(request_configuration=request_config)
        return folders

    async def get_messages(self, id):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            # Only request specific properties
            select=[
                "from",
                "isRead",
                "receivedDateTime",
                "subject",
                "toRecipients",
                "isDraft",
            ],
            # Get at most 1000 results
            top=1000,
            # Sort by received time, newest first
            orderby=["receivedDateTime DESC"],
        )
        request_config = (
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
        )

        messages = await self.user_client.me.mail_folders.by_mail_folder_id(
            id
        ).messages.get(request_configuration=request_config)
        return messages
