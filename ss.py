from benedict import benedict

vv = {"a":{"c": {"b":2}}}
vv = benedict(vv, keypath_separator='.')
print(vv["a"]["c"])
asd=1

print(vv.get('a.c.b'))
