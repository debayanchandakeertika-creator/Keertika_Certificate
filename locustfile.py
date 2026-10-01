from locust import HttpUser,task

class FirstLocust(HttpUser):
	@task
	def hello_world(self):
		self.client.get("/storeData?empid=a5c0c204d8735f18&uuid=03521025-4400-4453-9cdf-7b44994c9906&mac=03521025-4400-4453-9cdf-7b44994c9906&dis=0&rssi=0&lat=20.9223373&long=77.7465592&time=11 November 2022 06:21:23 pm")