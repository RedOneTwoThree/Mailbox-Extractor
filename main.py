import asyncio
import configparser
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph


async def main():
    print("Email extractor tool\n")

    # Load settings
    config = configparser.ConfigParser()
    config.read(["config.cfg", "config.dev.cfg"])
    azure_settings = config["azure"]

    graph: Graph = Graph(azure_settings)

    await greet_user(graph)

    choice = -1

    while choice != 0:
        print("Please choose one of the following options:")
        print("0. Exit")
        print("1. Display access token")
        print("2. List my inbox")
        print("3. List closed cases")

        try:
            choice = int(input())
        except ValueError:
            choice = -1

        try:
            if choice == 0:
                print("Goodbye...")
            elif choice == 1:
                await display_access_token(graph)
            elif choice == 2:
                await list_inbox(graph)
            elif choice == 3:
                await list_folder(graph)
            else:
                print("Invalid choice!\n")
        except ODataError as odata_error:
            print("Error:")
            if odata_error.error:
                print(odata_error.error.code, odata_error.error.message)


async def greet_user(graph: Graph):
    user = await graph.get_user()
    if user:
        print("Hello,", user.display_name)
        print("Email:", user.mail or user.user_principal_name, "\n")


async def display_access_token(graph: Graph):
    token = await graph.get_user_token()
    print("User token:", token, "\n")


async def list_inbox(graph: Graph):
    message_page = await graph.get_inbox()
    if message_page and message_page.value:
        # Output each message's details
        for message in message_page.value:
            print("Message:", message.subject)
            if message.from_ and message.from_.email_address:
                print("  From:", message.from_.email_address.name or "NONE")
            else:
                print("  From: NONE")
            print("  Status:", "Read" if message.is_read else "Unread")
            print("  Received:", message.received_date_time)

        # If @odata.nextLink is present
        more_available = message_page.odata_next_link is not None
        print("\nMore messages available?", more_available, "\n")


async def list_folder(graph: Graph):
    folder_page = await graph.get_folder()
    clientdict = {}
    index = 1
    if folder_page and folder_page.value:
        print("Clients")
        # Output each client
        for name in folder_page.value:
            print(str(index) + ". " + name.display_name)
            clientdict.update({name.display_name: name.id})
            index += 1
        print("Please select a client:")
    client = -1
    try:
        client = int(input())
    except ValueError:
        client = -1

    print("Client " + list(clientdict)[client - 1] + " selected")
    print(" ")
    id = list(clientdict.values())[client - 1]

    message_page = await graph.get_messages(id)
    if message_page and message_page.value:
        # Output each message's details
        x = 0
        for message in message_page.value:
            if message.is_draft == False:
                x += 1
                print("Subject:", message.subject)
                if message.to_recipients and message.to_recipients[0].email_address:
                    print(
                        "  To:", message.to_recipients[0].email_address.name or "NONE"
                    )
                else:
                    print("  To: NONE")
                if message.from_ and message.from_.email_address:
                    print("  From:", message.from_.email_address.name or "NONE")
                else:
                    print("  From: NONE")
                print("  Received:", message.received_date_time)
                print(" ")
        print("Total: ", x)


async def send_mail(graph: Graph):
    # TODO
    return


async def make_graph_call(graph: Graph):
    # TODO
    return


# Run main
asyncio.run(main())
