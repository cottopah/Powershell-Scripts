#Get total count of GPOs in AD

get-gpo -All -Domain tennant.tco.corp | measure


#Get  a list of GPOs in AD
#get-gpo -All -Domain tennant.tco.corp | select displayname, id