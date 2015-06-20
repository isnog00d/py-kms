import logging
import argparse
import binascii
import hashlib
import random
import socket
import SocketServer
import struct
import uuid
import rpcBind, rpcRequest
import sys
import os

from dcerpc import MSRPCHeader
from rpcBase import rpcBase

config = {}

def main():
	parser = argparse.ArgumentParser()
	parser.add_argument("ip", nargs="?", action="store", default="0.0.0.0", help="The IP address to listen on. The default is \"0.0.0.0\" (all interfaces).", type=str)
	parser.add_argument("port", nargs="?", action="store", default=1688, help="The network port to listen on. The default is \"1688\".", type=int)
	parser.add_argument("-e", "--epid", dest="epid", default=None, help="Use this flag to manually specify an ePID to use. If no ePID is specified, a random ePID will be generated.", type=str)
	parser.add_argument("-l", "--lcid", dest="lcid", default=1033, help="Use this flag to manually specify an LCID for use with randomly generated ePIDs. If an ePID is manually specified, this setting is ignored.", type=int)
	parser.add_argument("-c", "--client-count", dest="CurrentClientCount", default=26, help="Use this flag to specify the current client count. Default is 26. A number >25 is required to enable activation.", type=int)
	parser.add_argument("-a", "--activation-interval", dest="VLActivationInterval", default=120, help="Use this flag to specify the activation interval (in minutes). Default is 120 minutes (2 hours).", type=int)
	parser.add_argument("-r", "--renewal-interval", dest="VLRenewalInterval", default=1440 * 7, help="Use this flag to specify the renewal interval (in minutes). Default is 10080 minutes (7 days).", type=int)
	parser.add_argument("-v", "--loglevel", dest="loglevel", action="store", default="ERROR", choices=["CRITICAL", "ERROR", "WARNING", "INFO", "DEBUG"], help="set's Loglevel")
	parser.add_argument("-f", "--logfile", dest="logfile", action="store", default=os.path.dirname(os.path.abspath( __file__ )) + "/pykms.log", help="Logfile to write Output to")
	config.update(vars(parser.parse_args()))
	logging.basicConfig(filename=config['logfile'], level=config['loglevel'])
	server = SocketServer.TCPServer((config['ip'], config['port']), kmsServer)
	server.timeout = 5
	logging.info("TCP server listening at %s on port %d." % (config['ip'],config['port']))
	server.serve_forever()

class kmsServer(SocketServer.BaseRequestHandler):
	def setup(self):
		self.connection = self.request
		logging.info("Connection accepted: %s:%d" % (self.client_address[0],self.client_address[1]))

	def handle(self):
		while True:
			# self.request is the TCP socket connected to the client
			try:
				self.data = self.connection.recv(1024)
			except socket.error, e:
				if e[0] == 104:
					logging.error("Connection reset by peer.")
					break
				else:
					raise
			if self.data == '' or not self.data:
				logging.warn("No data received!")
				break
			# self.data = bytearray(self.data.strip())
			# logging.debug(binascii.b2a_hex(str(self.data)))
			packetType = MSRPCHeader(self.data)['type']
			if packetType == rpcBase.packetType['bindReq']:
				logging.info("RPC bind request received.")
				handler = rpcBind.handler(self.data, config)
			elif packetType == rpcBase.packetType['request']:
				logging.info("Received activation request.")
				handler = rpcRequest.handler(self.data, config)
			else:
				logging.error("Invalid RPC request type", packetType)
				break

			handler.populate()
			res = str(handler.getResponse())
			self.connection.send(res)

			if packetType == rpcBase.packetType['bindReq']:
				logging.info("RPC bind acknowledged.")
			elif packetType == rpcBase.packetType['request']:
				logging.info("Responded to activation request.")
				break

	def finish(self):
		self.connection.close()
		logging.info("Connection closed: %s:%d" % (self.client_address[0],self.client_address[1]))

if __name__ == "__main__":
	main()
