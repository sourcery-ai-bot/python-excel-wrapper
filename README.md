# Python Excel Wrapper


Le but du projet de se passer de l'utilisation de VBA, de faire un wrapper qui permet de lire et
écrire facilement et rapidement dans un excel. 

La problématique réside dans la lenteur de lecture de fichier Excel avec Python. Pour pallier à ce problème,
on converti dans un premier temps le fichier lu en csv. Cette opération est un peu lente alors on fait en sorte
d'enregistrer le hash du fichier dans un fichier .txt. Puis lors de la prochaine éxecution, on vérifiera d'abord
si le hash a changé comparé à l'ancien, et si ce n'est pas le cas, de récupérer le csv précédemment généré.
 