# Function-of-Ameripour-Hydrates-Model  
## The work of Sharareh Ameripour in "PREDICTION OF GAS-HYDRATE FORMATION CONDITIONS IN PRODUCTION AND SURFACE FACILITIES" is excelent. And is already in Excel.
So. Why im i in here adding more code?  
## I honestly cannot make an improvement of the theorical, practical and statistical solution she proposed. But, i also bealive that the user of Excel is looking for a robust, function focused solution. instead of looking for changing a sub or uploading it's values in another sheet.  

The original Sharareh sheet's are in a macro, not in a excel formula.  
So predicting a trending will be much harder with that approach.  
In here you will find the proposed equations, as a easy to use function.  
##How do you use the formula?  
=HydrateAmeripourS(TempInF,PressInPSI,LabelsOfComponents,ValuesOfComponents_MolPercentage)
Where...  
TempInF = Temperature in Farenheit  
PressInPSI = Pressure in PSI  
LabelsOfComponents = Name of the components beeing uploaded.  
ValuesOfComponents_MolPercentage = mole percent of the components beeing uploaded.  
## how do you expect me to use the formula without an example?  
for example you can use =HydrateAmeripourS(32,90,$B$8:$B$19,$C$8:$C$19)  
where the selected range can be something like this.  
![Columns](GitHub_1.png)  
or this.  
![Rows](GitHub_2.png)  
