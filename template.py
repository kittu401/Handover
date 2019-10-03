
def template(date_now,Present_shift,Present_shift_Mems,ShiftEm,Next_Shift_Mems,Avaya_ALert_count,Data_Alert_Count,
             Avaya_Tkt_count_And_Proirities,Data_Tkt_count_And_Proirities,Total_ticket_count,Priority_Grand_Total,
             grand_total_sla,Sla_Percentages):
  Nectar_cont = round((Avaya_Tkt_count_And_Proirities[0]/Avaya_ALert_count[1])*100)
  AlarmTrac_cont=round((Avaya_Tkt_count_And_Proirities[1]/Avaya_ALert_count[0])*100)
  Iris_Cont = round((Avaya_Tkt_count_And_Proirities[2]/Avaya_ALert_count[2])*100)

  OnBoard_cont =round((Data_Tkt_count_And_Proirities[1]/Data_Alert_Count[0])*100)
  BCG_cont = round((Data_Tkt_count_And_Proirities[0]/Data_Alert_Count[1])*100)
  Alerts_Grand_Total=Avaya_ALert_count[0] + Avaya_ALert_count[1] + Avaya_ALert_count[2] + Data_Alert_Count[0] + Data_Alert_Count[1]
  Tickets_Grand_Total = Avaya_Tkt_count_And_Proirities[0]+Avaya_Tkt_count_And_Proirities[1]+Avaya_Tkt_count_And_Proirities[2]\
                        +Data_Tkt_count_And_Proirities[0]+Data_Tkt_count_And_Proirities[1]

  Total_cont = round((Tickets_Grand_Total/Alerts_Grand_Total)*100)

  with open("Shift_handover.html", "w") as my_file:
    my_file.write('''<html>

<head>
	<!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<!-- Optional theme -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">

<title> Handover</title>
</head>


<body>
	<div class="conatiner">
		<center>
		<h1>EVENT MANAGEMENT HANDOVER</h1>
	
		<br>	
		<br>
	<table class="table table-striped">
		<tr>
		  <th>Date</th>''')
    my_file.write("<th>%s</th>" %date_now)

    my_file.write("<th>%s</th>" %Present_shift)
    my_file.write('''<th>Handover details</th>
        <th>SHIFT EM</th>
		</tr>
		<tr>
			<th> NOC Analysts/EM:</th>''')
    my_file.write("<td colspan=3>%s</td>" %Present_shift_Mems)
    my_file.write("<td>%s</td>"%ShiftEm)
    my_file.write('''</tr>
		<tr>
			<th><b>Handover to NOC Analysts Next Shift/EM-Next Shift</b></th>''')
    my_file.write("<td colspan=3>%s</td>" %Next_Shift_Mems)

    my_file.write(''' </tr>
	</table>
	</center>''')
  
    my_file.write('''
    
	<br><br><br>
	<div class="row">

  <div class="table-responsive col-md-6">
  	<center><h1>ALERT DASHBOARD TABLE</h1></center>
      <table class="table table-striped">
      	<tr>
			<th>Alert Dashboard</th>
			<th>Source</th>
			<th>Alert Received</th>
			<th>Ticket Created</th>
			<th>Contb %</th>
		</tr>
		<tr>
			<td>Warwick</td>
			<td>Nimsoft</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
		</tr>
		<tr>
			<td>Exeter</td>
			<td>Nimsoft</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
		</tr>
		<tr>
			<td>Nectar</td>
			<td>Nectar</td>''')
    my_file.write("<td>%s</td>" %Avaya_ALert_count[1])
    my_file.write("<td>%s</td>" %Avaya_Tkt_count_And_Proirities[0])
    my_file.write("<td>%s</td>" %Nectar_cont)
    my_file.write('''</tr>
		<tr>
			<td>Alarmtrac</td>
			<td>Alarmtrac</td>''')
    my_file.write("<td>%s</td>"%Avaya_ALert_count[0])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[1])
    my_file.write("<td>%s</td>"%AlarmTrac_cont)
    my_file.write('''</tr>
		<tr>
			<td>IRIS</td>
			<td>nortel</td>''')
    my_file.write("<td>%s</td>"%Avaya_ALert_count[2])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[2])
    my_file.write("<td>%s</td>"%Iris_Cont)
    my_file.write('''</tr>
		<tr>
			<td>On Boarding</td>
			<td>Vistara</td>''')
    my_file.write("<td>%s</td>"%Data_Alert_Count[0])
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[1])
    my_file.write("<td>%s</td>"%OnBoard_cont)
    my_file.write('''</tr>
		<tr>
			<td>BCG</td>
			<td>Vistara</td>''')
    my_file.write("<td>%s</td>"%Data_Alert_Count[1])
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[0])
    my_file.write("<td>%s</td>"%BCG_cont)
    my_file.write('''</tr>
		<tr>
			<td colspan=2>Grand total</td>''')

    my_file.write("<td>%s</td>"%Alerts_Grand_Total)
    my_file.write("<td>%s</td>"%Tickets_Grand_Total)
    my_file.write("<td>%s</td>"%Total_cont)
    my_file.write('''</tr>
    

      </table>
  </div>
<!---------------SLA DASHBOARD TABLE-------------->
<center><h1>SLA DASHBOARD TABLE</h1></center>
  		<div class="table-responsive col-md-6">
  			
			<table class="table table-striped">
				<tr>
					<th rowspan=2>SLA Dashboard</th>
					<th>P1%</th>
					<th>P2%</th>
					<th>P3%</th>
					<th>P4%</th>
					<th rowspan=2>SLA %age</th>
				</tr>

				<tr>
					<th>critical</th>
					<th>High</th>
					<th>Medium</th>
					<th>Low</th>
					
				</tr>

				<tr>
					<td>Warwick</td>
					<td>--</td>
					<td>--</td>
					<td>--</td>
					<td>--</td>
					<td>--</td>
				</tr>

				<tr>
					<td>Exeter</td>
					<td>--</td>
					<td>--</td>
					<td>--</td>
					<td>--</td>
					<td>--</td>
				</tr>

				<tr>
					<td>Nectar</td>''')
    my_file.write("<td>%s</td>"%Sla_Percentages[0][0])
    my_file.write("<td>%s</td>" % Sla_Percentages[0][1])
    my_file.write("<td>%s</td>" % Sla_Percentages[0][2])
    my_file.write("<td>%s</td>" % Sla_Percentages[0][3])
    try:
      my_file.write("<td>%s</td>" %round((Avaya_Tkt_count_And_Proirities[4][0] + Avaya_Tkt_count_And_Proirities[4][2] +
                                     Avaya_Tkt_count_And_Proirities[4][4]
                                     + Avaya_Tkt_count_And_Proirities[4][6])/(Total_ticket_count[0])*100))
    except Exception:
      my_file.write(("<td>%s</td>" % 0))

    my_file.write('''</tr>
				<tr>
					<td>Alarmtrac</td>''')
    my_file.write("<td>%s</td>" % Sla_Percentages[1][0])
    my_file.write("<td>%s</td>" % Sla_Percentages[1][1])
    my_file.write("<td>%s</td>" % Sla_Percentages[1][2])
    my_file.write("<td>%s</td>" % Sla_Percentages[1][3])
    try:
      my_file.write("<td>%s</td>" % round((Avaya_Tkt_count_And_Proirities[3][0] + Avaya_Tkt_count_And_Proirities[3][2] +
                                     Avaya_Tkt_count_And_Proirities[3][4]
                                     + Avaya_Tkt_count_And_Proirities[3][6]) /(Total_ticket_count[1])*100))
    except Exception:
      my_file.write(("<td>%s</td>" % 0))

    my_file.write('''</tr>
				<tr>
					<td>IRIS</td>''')
    my_file.write("<td>%s</td>" % Sla_Percentages[2][0])
    my_file.write("<td>%s</td>" % Sla_Percentages[2][1])
    my_file.write("<td>%s</td>" % Sla_Percentages[2][2])
    my_file.write("<td>%s</td>" % Sla_Percentages[2][3])
    try:
      my_file.write("<td>%s</td>" % round((Avaya_Tkt_count_And_Proirities[5][0]+Avaya_Tkt_count_And_Proirities[5][2]
                                      +Avaya_Tkt_count_And_Proirities[5][4]
                                   +Avaya_Tkt_count_And_Proirities[5][6]) / (Total_ticket_count[2])*100))
    except Exception:
      my_file.write(("<td>%s</td>" % 0))

    my_file.write('''</tr>
				<tr>
					<td>On Boarding</td>''')
    my_file.write("<td>%s</td>" % Sla_Percentages[3][0])
    my_file.write("<td>%s</td>" % Sla_Percentages[3][1])
    my_file.write("<td>%s</td>" % Sla_Percentages[3][2])
    my_file.write("<td>%s</td>" % Sla_Percentages[3][3])
    try:
      my_file.write("<td>%s</td>" % round((Data_Tkt_count_And_Proirities[3][0] + Data_Tkt_count_And_Proirities[3][2] +
                                     Data_Tkt_count_And_Proirities[3][4]
                                     + Data_Tkt_count_And_Proirities[3][6])/(Total_ticket_count[3])*100))
    except Exception:
      my_file.write(("<td>%s</td>" % 0))

    my_file.write('''</tr>
				<tr>
					<td>BCG</td>''')
    my_file.write("<td>%s</td>" % Sla_Percentages[4][0])
    my_file.write("<td>%s</td>" % Sla_Percentages[4][1])
    my_file.write("<td>%s</td>" % Sla_Percentages[4][2])
    my_file.write("<td>%s</td>" % Sla_Percentages[4][3])
    try:
      my_file.write("<td>%s</td>" % round((Data_Tkt_count_And_Proirities[2][0] + Data_Tkt_count_And_Proirities[2][2] +
                                     Data_Tkt_count_And_Proirities[2][4]
                                     + Data_Tkt_count_And_Proirities[2][6]) / (Total_ticket_count[4])*100))
    except Exception:
      my_file.write(("<td>%s</td>" %0))



    my_file.write('''</tr>
				<tr>
					<td>Grand total</td>''')
    my_file.write("<td>%s</td>" % round((Data_Tkt_count_And_Proirities[2][0] + Data_Tkt_count_And_Proirities[3][0] +
                                     Avaya_Tkt_count_And_Proirities[3][0] +
                                     Avaya_Tkt_count_And_Proirities[4][0] + Avaya_Tkt_count_And_Proirities[5][0]) /
                                         (Priority_Grand_Total[0]) * 100))
    my_file.write("<td>%s</td>" % round((Data_Tkt_count_And_Proirities[2][2] + Data_Tkt_count_And_Proirities[3][2] +
                                     Avaya_Tkt_count_And_Proirities[3][2] +
                                     Avaya_Tkt_count_And_Proirities[4][2] + Avaya_Tkt_count_And_Proirities[5][2]) /
                                          (Priority_Grand_Total[1]) * 100))
    my_file.write("<td>%s</td>" % round((Data_Tkt_count_And_Proirities[2][4] + Data_Tkt_count_And_Proirities[3][4] +
                                     Avaya_Tkt_count_And_Proirities[3][4] +
                                     Avaya_Tkt_count_And_Proirities[4][4] + Avaya_Tkt_count_And_Proirities[5][4]) /
                                    (Priority_Grand_Total[2]) * 100))
    my_file.write("<td>%s</td>" % round((Data_Tkt_count_And_Proirities[2][4] + Data_Tkt_count_And_Proirities[3][4] +
                                     Avaya_Tkt_count_And_Proirities[3][4] +
                                     Avaya_Tkt_count_And_Proirities[4][4] + Avaya_Tkt_count_And_Proirities[5][4]) /
                                    (Priority_Grand_Total[2]) * 100))
    my_file.write("<td>%s</td>"%round((grand_total_sla[0])/(grand_total_sla[1]+grand_total_sla[0])*100))
    my_file.write('''</tr>

			</table>
		</div>

<!--------------Priority Wise Dashboard -------------------->

<div class="conatiner">
	<center><h1>Priority Wise Dashboard</h1></center>
	<table class="table table-striped">
		<tr>
			<th rowspan=2>PriorityWise Dashboard</th>
			<th >P1</th>
			<th >P2</th>
			<th >P3</th>
			<th >P4</th>
			<th>Grand Total</th>
		</tr>
		<tr>
			<th>Critical</th>
			<th>High</th>
			<th>Medium</th>
			<th>Normal/Low</th>
			<th>ALL POD's</th>
		</tr>

		<table class="table table-striped">
			<tr>
			<th>Legends</th>
			<th>INSLA</th>
			<th>OSLA</th>
			<th>Total</th>

			<th>INSLA</th>
			<th>OSLA</th>
			<th>Total</th>

			<th>INSLA</th>
			<th>OSLA</th>
			<th>Total</th>

			<th>INSLA</th>
			<th>OSLA</th>
			<th>Total</th>

			<th>INSLA</th>
			<th>OSLA</th>
			<th>Total</th>
		</tr>
		
		<tr>
			<td>Warwick</td>
            <td>--</td>
            <td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			
		</tr>	
		<tr>
			<td>Exeter</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			<td>--</td>
			
		</tr>	
		<tr>
			<td>Nectar</td>''')
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][0])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][1])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[4][0]+Avaya_Tkt_count_And_Proirities[4][1]))
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][2])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][3])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[4][2]+Avaya_Tkt_count_And_Proirities[4][3]))
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][4])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][5])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[4][4]+Avaya_Tkt_count_And_Proirities[4][5]))
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][6])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[4][7])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[4][6]+Avaya_Tkt_count_And_Proirities[4][7]))
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[4][0] + Avaya_Tkt_count_And_Proirities[4][2] +
                                   Avaya_Tkt_count_And_Proirities[4][4]
                                   + Avaya_Tkt_count_And_Proirities[4][6]))
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[4][1] + Avaya_Tkt_count_And_Proirities[4][3] +
                                   Avaya_Tkt_count_And_Proirities[4][5]
                                   + Avaya_Tkt_count_And_Proirities[4][7]))
    my_file.write("<td>%s</td>"%Total_ticket_count[0])
			
    my_file.write('''</tr>
		<tr>
			<td>Alarmtrac</td>''')
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][0])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][1])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[3][0]+Avaya_Tkt_count_And_Proirities[3][1]))
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][2])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][3])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[3][2]+Avaya_Tkt_count_And_Proirities[3][3]))
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][4])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][5])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[3][4]+Avaya_Tkt_count_And_Proirities[3][5]))
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][6])
    my_file.write("<td>%s</td>"%Avaya_Tkt_count_And_Proirities[3][7])
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[3][6]+Avaya_Tkt_count_And_Proirities[3][7]))
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[3][0] + Avaya_Tkt_count_And_Proirities[3][2] +
                                   Avaya_Tkt_count_And_Proirities[3][4]
                                   + Avaya_Tkt_count_And_Proirities[3][6]))
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[3][1] + Avaya_Tkt_count_And_Proirities[3][3] +
                                   Avaya_Tkt_count_And_Proirities[3][5]
                                   + Avaya_Tkt_count_And_Proirities[3][7]))
    my_file.write("<td>%s</td>"%Total_ticket_count[1])
			
    my_file.write('''</tr>
		<tr>
			<td>IRIS</td>''')
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][0])
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][1])
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[5][0] + Avaya_Tkt_count_And_Proirities[5][1]))
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][2])
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][3])
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[5][2] + Avaya_Tkt_count_And_Proirities[5][3]))
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][4])
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][5])
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[5][4] + Avaya_Tkt_count_And_Proirities[5][5]))
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][6])
    my_file.write("<td>%s</td>" % Avaya_Tkt_count_And_Proirities[5][7])
    my_file.write("<td>%s</td>" % (Avaya_Tkt_count_And_Proirities[5][6] + Avaya_Tkt_count_And_Proirities[5][7]))
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[5][0]+Avaya_Tkt_count_And_Proirities[5][2]+Avaya_Tkt_count_And_Proirities[5][4]
                                 +Avaya_Tkt_count_And_Proirities[5][6]))
    my_file.write("<td>%s</td>"%(Avaya_Tkt_count_And_Proirities[5][1]+Avaya_Tkt_count_And_Proirities[5][3]+Avaya_Tkt_count_And_Proirities[5][5]
                                 +Avaya_Tkt_count_And_Proirities[5][7]))
    my_file.write("<td>%s</td>"%Total_ticket_count[2])
    my_file.write('''</tr>
		<tr>
			<td>On Boarding</td>''')
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][0])
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][1])
    my_file.write("<td>%s</td>"%(Data_Tkt_count_And_Proirities[3][0]+Data_Tkt_count_And_Proirities[3][1]))
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][2])
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][3])
    my_file.write("<td>%s</td>"%(Data_Tkt_count_And_Proirities[3][2]+Data_Tkt_count_And_Proirities[3][3]))
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][4])
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][5])
    my_file.write("<td>%s</td>"%(Data_Tkt_count_And_Proirities[3][4]+Data_Tkt_count_And_Proirities[3][5]))
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][6])
    my_file.write("<td>%s</td>"%Data_Tkt_count_And_Proirities[3][7])
    my_file.write("<td>%s</td>"%(Data_Tkt_count_And_Proirities[3][6]+Data_Tkt_count_And_Proirities[3][7]))
    my_file.write("<td>%s</td>"%(Data_Tkt_count_And_Proirities[3][0]+Data_Tkt_count_And_Proirities[3][2]+Data_Tkt_count_And_Proirities[3][4]
                                 +Data_Tkt_count_And_Proirities[3][6]))
    my_file.write("<td>%s</td>"%(Data_Tkt_count_And_Proirities[3][1]+Data_Tkt_count_And_Proirities[3][3]+Data_Tkt_count_And_Proirities[3][5]
                                 +Data_Tkt_count_And_Proirities[3][7]))
    my_file.write("<td>%s</td>"%Total_ticket_count[3])
    my_file.write('''</tr>
		<tr>
			<td>BCG</td>''')
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][0])
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][1])
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][0] + Data_Tkt_count_And_Proirities[2][1]))
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][2])
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][3])
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][2] + Data_Tkt_count_And_Proirities[2][3]))
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][4])
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][5])
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][4] + Data_Tkt_count_And_Proirities[2][5]))
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][6])
    my_file.write("<td>%s</td>" % Data_Tkt_count_And_Proirities[2][7])
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][6] + Data_Tkt_count_And_Proirities[2][7]))
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][0] + Data_Tkt_count_And_Proirities[2][2] +
                                   Data_Tkt_count_And_Proirities[2][4]
                                   + Data_Tkt_count_And_Proirities[2][6]))
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][1] + Data_Tkt_count_And_Proirities[2][3] +
                                   Data_Tkt_count_And_Proirities[2][5]
                                   + Data_Tkt_count_And_Proirities[2][7]))
    my_file.write("<td>%s</td>"%Total_ticket_count[4])

    my_file.write('''</tr>
		<tr>
            <td>Grand total</td>''')
    my_file.write("<td>%s</td>"%(Data_Tkt_count_And_Proirities[2][0]+Data_Tkt_count_And_Proirities[3][0]+Avaya_Tkt_count_And_Proirities[3][0]+
                                 Avaya_Tkt_count_And_Proirities[4][0]+Avaya_Tkt_count_And_Proirities[5][0]))
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][1] + Data_Tkt_count_And_Proirities[3][1] +
                                   Avaya_Tkt_count_And_Proirities[3][1] +
                                   Avaya_Tkt_count_And_Proirities[4][1] + Avaya_Tkt_count_And_Proirities[5][1]))
    my_file.write("<td>%s</td>" %Priority_Grand_Total[0])

    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][2] + Data_Tkt_count_And_Proirities[3][2] +
                                   Avaya_Tkt_count_And_Proirities[3][2] +
                                   Avaya_Tkt_count_And_Proirities[4][2] + Avaya_Tkt_count_And_Proirities[5][2]))
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][3] + Data_Tkt_count_And_Proirities[3][3] +
                                   Avaya_Tkt_count_And_Proirities[3][3] +
                                   Avaya_Tkt_count_And_Proirities[4][3] + Avaya_Tkt_count_And_Proirities[5][3]))
    my_file.write("<td>%s</td>" %Priority_Grand_Total[1])
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][4] + Data_Tkt_count_And_Proirities[3][4] +
                                   Avaya_Tkt_count_And_Proirities[3][4] +
                                   Avaya_Tkt_count_And_Proirities[4][4] + Avaya_Tkt_count_And_Proirities[5][4]))
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][5] + Data_Tkt_count_And_Proirities[3][5] +
                                   Avaya_Tkt_count_And_Proirities[3][5] +
                                   Avaya_Tkt_count_And_Proirities[4][5] + Avaya_Tkt_count_And_Proirities[5][5]))
    my_file.write("<td>%s</td>" % Priority_Grand_Total[2])
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][6] + Data_Tkt_count_And_Proirities[3][6] +
                                   Avaya_Tkt_count_And_Proirities[3][6] +
                                   Avaya_Tkt_count_And_Proirities[4][6] + Avaya_Tkt_count_And_Proirities[5][6]))
    my_file.write("<td>%s</td>" % (Data_Tkt_count_And_Proirities[2][7] + Data_Tkt_count_And_Proirities[3][7] +
                                   Avaya_Tkt_count_And_Proirities[3][7] +
                                   Avaya_Tkt_count_And_Proirities[4][7] + Avaya_Tkt_count_And_Proirities[5][7]))
    my_file.write("<td>%s</td>" % Priority_Grand_Total[3])
    my_file.write("<td>%s</td>"%grand_total_sla[0])
    my_file.write("<td>%s</td>" % grand_total_sla[1])
    my_file.write("<td>%s</td>" % (grand_total_sla[1]+grand_total_sla[0]))

    my_file.write('''</tr>
		</table>
	</table>

</div>
<br><br>
</div class=conatiner>
	<table class="table table-striped">
		<tr>
			<th colspan=4>Important Update **</th>
		</tr>
		<tr>
			<td colspan=4>Write Updates Here</td>
		</tr>
	<!-- Latest compiled and minified JavaScript -->
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
</body>
    
   </html> 
    ''')

