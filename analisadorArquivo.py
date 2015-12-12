from pyspark import SparkConf, SparkContext

class AnalisadorArquivo(object):

	conf = SparkConf
	sc = SparkContext

	def __init__(self):
		print 'Iniciando o analisador de arquivos'

	def loadConfig(self):
		conf = (SparkConf().setMaster("local").setAppName("Analisador").set("SparkConfrk.executor.memory", "1g"))
		return conf
		
	def loadContext(self, conf):
		sc = SparkContext(conf = conf)
		return sc

	def loadFile(self, sc, pathFile):
		loadFile = sc.textFile(pathFile)
		return loadFile

	def countWords(self, rdd):
		return rdd.count()

	def filter(self, rdd, palavra):
		filtro = rdd.filter(lambda line: palavra in line)
		
	def mapReduceSave(self, rdd, pathOutFile):
		result = rdd.flatMap(lambda line: line.split(' ')).map(lambda word: (word, 1)).reduceByKey(lambda a,b: a+b)
		result.saveAsTextFile(pathOutFile)
		
if __name__ == "__main__":
	analise = AnalisadorArquivo()
	confg = analise.loadConfig()
	sparkContext = analise.loadContext(confg)
	rdd = analise.loadFile(sparkContext,'/Users/Rolemberg/Documents/python/spark/testeSentimental.txt')
	
	analise.mapReduceSave(rdd,'/Users/Rolemberg/Documents/python/spark/testeout1')
	# analise.filter(rdd,'info')

	# totalPalavras = analise.countWords(rdd)
	# print 'Total de palavras encontradas no seu arquivo', totalPalavras