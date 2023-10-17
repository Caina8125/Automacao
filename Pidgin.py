import xmpp

def financeiroDemo(texto):
    pidgin_Id= "alertas.xmpp@chat.amhp.local"
    senha    = "!!87316812"
    dev       = "lucas.paz@chat.amhp.local"

    jid = xmpp.protocol.JID(pidgin_Id)
    connection = xmpp.Client(server=jid.getDomain())
    connection.connect()
    connection.auth(user=jid.getNode(), password=senha, resource=jid.getResource())
    connection.send(xmpp.protocol.Message(to=dev, body=texto))

def notaFiscal(texto):
    count     = 0
    pidgin_Id = "notafiscal.bot@chat.amhp.local"
    senha     = "!!87316812"
    users     = ["lucas.timoteo@chat.amhp.local"
                 "brenda.moura@chat.amhp.local",
                 "isabella.souza@chat.amhp.local"]

    jid = xmpp.protocol.JID(pidgin_Id)
    connection = xmpp.Client(server=jid.getDomain())
    connection.connect()
    connection.auth(user=jid.getNode(), password=senha, resource=jid.getResource())
    for i in range(3):
        connection.send(xmpp.protocol.Message(to=users[count], body=texto))
        count += 1