import xmpp

pidgin_Id= "alertas.xmpp@chat.amhp.local"
senha    = "!!87316812"
msg      = "Bot_CTI Funcionando"

dev     = "lucas.paz@chat.amhp.local"

def main(texto):
    jid = xmpp.protocol.JID(pidgin_Id)
    connection = xmpp.Client(server=jid.getDomain())
    connection.connect()
    connection.auth(user=jid.getNode(), password=senha, resource=jid.getResource())
    connection.send(xmpp.protocol.Message(to=dev, body=texto))

if __name__ == "__main__":
    main()