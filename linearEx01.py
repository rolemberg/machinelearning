from pyspark.mllib.classification import LogisticRegressionWithSGD
#from pyspark.mllib.regression import LabeledPoint


class Linear(object):

	def __init__(self):
		print 'Teste linear'
	# Load and parse the data
	#def parsePoint(line):
 	#	values = [float(x) for x in line.split(' ')]
    #	return LabeledPoint(values[0], values[1:])


if __name__ == "__main__":
	print 'teste'
	# data = sc.textFile("/Users/Rolemberg/Documents/python/spark/machinelearning/sample_svm_data.txt")
	# parsedData = data.map(parsePoint)

	# Build the model
	# model = SVMWithSGD.train(parsedData, iterations=100)

	# Evaluating the model on training data
	# labelsAndPreds = parsedData.map(lambda p: (p.label, model.predict(p.features)))
	# trainErr = labelsAndPreds.filter(lambda (v, p): v != p).count() / float(parsedData.count())
	# print("Training Error = " + str(trainErr))

	# Save and load model
	# model.save(sc, "myModelPath")
	# sameModel = SVMModel.load(sc, "myModelPath")

	