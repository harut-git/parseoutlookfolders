import sys
import json

from gevent.wsgi import WSGIServer

print sys.version_info


def application(environ, start_response):
    start_response('200 OK', [('Content-Type', 'text/html')])
    # start_response('200 OK', [('Access-Control-Allow-Origin','*')])
    request = environ['wsgi.input'].read()
    print "request:", request
    return "done"


print 12313
WSGIServer(('', 9092), application).serve_forever()
