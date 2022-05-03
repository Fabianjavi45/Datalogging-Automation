class Date():

    def __init__(this,day,month,year):
        this.day=day
        this.month=month
        this.year=year

    #Gets the week 
    def getWeek(this):
        if(this.month==1 or this.month==3 or this.month==5 or this.month==7 
        or this.month==8 or this.month==10 or this.month==12):
            if(this.day>25):
                newDay=6-(31-this.day)
                newMonth=this.month+1
                newWeek= Date(newDay,newMonth,this.year)
                return newWeek
        if(this.month==4 or this.month==6 or this.month==9 or this.month==9 or this.month==11):
            if(this.day>24):
                newDay=6-(30-this.day)
                newMonth=this.month+1
                newWeek= Date(newDay,newMonth,this.year)
                return newWeek
            
        if(this.month==2):
            if(this.day>22):
                newDay=6-(28-this.day)
                newMonth=this.month+1
                newWeek= Date(newDay,newMonth,this.year)
                return newWeek
        newDay=this.day+6
        return Date(newDay,this.month,this.year)

    def getendWeek(this,weekDay):
        if(weekDay=="M"):
            weekDay=2
        elif(weekDay=="W"):
            weekDay=3
        elif(weekDay=="J"):
            weekDay=4
        elif(weekDay=="V"):
            weekDay=5
        
        if(this.month==1 or this.month==3 or this.month==5 or this.month==7 
        or this.month==8 or this.month==10 or this.month==12):
            if(this.day>25):
                newDay=(7-weekDay)+this.day
                if(newDay>31):
                    newDay=(7-weekDay)-(31-this.day)
                    newMonth=this.month+1
                    newWeek= Date(newDay,newMonth,this.year)
                    return newWeek
                newWeek= Date(newDay,this.month,this.year)
                return newWeek
            else:
                newDay=(7-weekDay)+this.day
                endWeek= Date(newDay,this.month,this.year)
                return endWeek
        if(this.month==4 or this.month==6 or this.month==9 or this.month==9 or this.month==11):
            if(this.day>24):
                newDay=(7-weekDay)+this.day
                if(newDay>30):
                    newDay=(7-weekDay)-(30-this.day)
                    newMonth=this.month+1
                    newWeek= Date(newDay,newMonth,this.year)
                    return newWeek
                newWeek= Date(newDay,this.month,this.year)
                return newWeek
            else:
                newDay=(7-weekDay)+this.day
                endWeek= Date(newDay,this.month,this.year)
                return endWeek    
        if(this.month==2):
            if(this.day>22):
                newDay=(7-weekDay)+this.day
                if(newDay>28):
                    newDay=(7-weekDay)-(28-this.day)
                    newMonth=this.month+1
                    newWeek= Date(newDay,newMonth,this.year)
                    return newWeek
                newWeek= Date(newDay,this.month,this.year)
                return newWeek
            else:
                newDay=(7-weekDay)+this.day
                endWeek= Date(newDay,this.month,this.year)
                return endWeek
    def getdayNumber(this,weekDay):
        if(weekDay=="M"):
            weekDay=2
        if(weekDay=="W"):
            weekDay=3
        if(weekDay=="J"):
            weekDay=4
        if(weekDay=="V"):
            weekDay=5
        return weekDay