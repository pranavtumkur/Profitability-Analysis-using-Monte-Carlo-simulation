# Profitability-Analysis-using-Monte-Carlo-simulation

**To access this project workbook you need to** -> [Click Here!](https://github.com/pranavtumkur/Profitability-Analysis-using-Monte-Carlo-simulation/raw/main/Profitability-Analysis-Tumkur-Monte%20Carlo.xlsm)

In this project, we determine if setting up a business would be profitable 15 years down the line, subject to 9 inputs, each associated with a fixed uncertainty. Following are the inputs which we consider and their specifications-

![Screenshot (43)](https://user-images.githubusercontent.com/65482013/103472748-17e26900-4db7-11eb-8ce0-cf0141c7110f.png)

To calculate the Cash Flow (CF) for a particular year, we will use the formula-

CF= (1-t)(S-C) + D - C<sub>TDC</sub> - C<sub>WC</sub> - C<sub>land</sub> - C<sub>startup</sub> - C<sub>royalty</sub> + V<sub>salvage</sub> (+ C<sub>WC</sub>)

Using Monte Carlo simulation, we iterate through a fixed number of scenarios, to determine what percent of scenarios result in a positive NPV (Net Present Value)-

![Screenshot (48)](https://user-images.githubusercontent.com/65482013/103472861-76f4ad80-4db8-11eb-8674-c7831bc374b2.png)

The user can also change the parameters using the userform displayed after clicking on 'Run Simulation' on the excel sheet-

![Screenshot (44)](https://user-images.githubusercontent.com/65482013/103472821-0057b000-4db8-11eb-810c-5ee26b1bcb12.png)

We then plot a histogram to get a 1-stop view of the distribution of the NPV as a consolidation of all the scenarios-

![Screenshot (49)](https://user-images.githubusercontent.com/65482013/103472857-63e1dd80-4db8-11eb-8f3c-2832b8f62cca.png)

