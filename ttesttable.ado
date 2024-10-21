
capture program drop ttest_table
program define ttest_table
	
	syntax varlist(max=1 fv) [if] [in], ///
		FILEname(string)	///	*  these are required options
		[STATs(string) ///	 these are optional options
		DECimals(string) title(string) NOTEs(string) ///
		font(string) FONTSize(string) notesize(string) ///
		txtindent(string) SINGleborder  ///
		group(varlist min=1 max=1) VARNAMEs noborder]
		
  putexcel set "`filename'.xlsx", sheet("TTest Table") replace
  
  
	local dv "`varlist'"   // variable name of your DV
    local iv "`group'"  // variable name of your IV		
	
*Set fonts/sizes for table
if "`font'" == "" {
	local fonttype = "timesnewroman"
	}

else {
	local fonttype = `" `font' "'
	}

if "`fontsize'" == "" {
	local fsize = "11"
	}

else {
	local fsize = `" `fontsize' "'
	}

if "`notesize'" == "" {
	local notefsize = "9"
	}
  
else {
	local notefsize = `" `notesize' "'
	}
	
	
** DO NOT EDIT THE FOLLOWING **  
 
 lab list `:val lab `iv''
 return list
	
 local ivlow  `r(min)' // the lower number assigned to the categories of your IV 
 local ivhigh `r(max)' // the higher number assigned to the categories of your IV 
   

local var_lblIV : variable label `iv'  
local var_lblDV : variable label `dv'  
local val_lbl0: label (`iv') `ivlow'
local val_lbl1: label (`iv') `ivhigh'

qui ttest `dv', by(`iv')
return list

local N_tot = `r(N_1)' + `r(N_2)'


if "`title'" == "" {
	qui putexcel A1:G1 = "Table #: Mean Differences in `var_lblDV' by `var_lblIV' (N=`N_tot')",  ///
				merge left border(bottom, `lineset') ///
				font("`fonttype'", "`fsize'") bold
	}

	else {
	local titletext = `" `title' "'
	qui putexcel A1:G1 = "`titletext'",  ///
				merge left border(bottom, `lineset') ///
				font("`fonttype'", "`fsize'") bold
	}


putexcel B2 = "N"
putexcel C2 = "Mean"
putexcel D2 = "S.E."
putexcel E2 = "Diff"
putexcel F2 = "t-value"
putexcel G2 = "two-tailed p-value"

putexcel A3 = "`val_lbl0'"
putexcel A4 = "`val_lbl1'"

putexcel B3 = `r(N_1)'
putexcel B4 = `r(N_2)'

local se_1 = `r(sd_1)'/sqrt(`r(N_1)')
local se_2 = `r(sd_2)'/sqrt(`r(N_2)')

putexcel C3 = `r(mu_1)', nformat(number_d2)
putexcel C4 = `r(mu_2)', nformat(number_d2)

putexcel D3 = `se_1', nformat(number_d2)
putexcel D4 = `se_2', nformat(number_d2)

putexcel E3:E4 = `r(mu_1)'-`r(mu_2)', nformat(number_d2) merge

putexcel F3:F4 = `r(t)', nformat(number_d2) merge
  
putexcel G3:G4 = `r(p)', nformat(number_d2) merge

 
end
