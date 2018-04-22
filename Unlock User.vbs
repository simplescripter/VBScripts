Set objUser = GetObject _
  ("LDAP://cn=homers,ou=safety,dc=nwtraders,dc=msft")
objUser.IsAccountLocked = False
objUser.SetInfo