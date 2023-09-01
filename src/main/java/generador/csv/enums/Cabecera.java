package generador.csv.enums;

import lombok.Getter;

@Getter
public enum Cabecera {

    C1("Name",0)  ,
    C2("Given Name",1)  ,
    C3("Additional Name",2)  ,
    C4("Family Name",3)  ,
    C5("Yomi Name",4)  ,
    C6("Given Name Yomi",5)  ,
    C7("Additional Name Yomi",6)  ,
    C8("Family Name Yomi",7)  ,
    C9("Name Prefix",8)  ,
    C10("Name Suffix",9)  ,
    C11("Initials",10)  ,
    C12("Nickname",11)  ,
    C13("Short Name",12)  ,
    C14("Maiden Name",13)  ,
    C15("Birthday",14)  ,
    C16("Gender",15)  ,
    C17("Location",16)  ,
    C18("Billing Information",17)  ,
    C19("Directory Server",18)  ,
    C20("Mileage",19)  ,
    C21("Occupation",20)  ,
    C22("Hobby",21)  ,
    C23("Sensitivity",22)  ,
    C24("Priority",23)  ,
    C25("Subject",24)  ,
    C26("Notes",25)  ,
    C27("Language",26)  ,
    C28("Photo",27)  ,
    C29("Group Membership",28)  ,
    C30("E-mail 1 - Type",29)  ,
    C31("E-mail 1 - Value",30)  ,
    C32("Phone 1 - Type",31)  ,
    C33("Phone 1 - Value",32)  ,
    C34("Organization 1 - Type",33)  ,
    C35("Organization 1 - Name",34)  ,
    C36("Organization 1 - Yomi Name",35)  ,
    C37("Organization 1 - Title",36)  ,
    C38("Organization 1 - Department",37)  ,
    C39("Organization 1 - Symbol",38)  ,
    C40("Organization 1 - Location",39)  ,
    C41("Organization 1 - Job Description",40) ;

    final String nombre ;
    final int posicion ;

    Cabecera(String nombre, int posicion) {
        this.nombre = nombre;
        this.posicion = posicion;
    }
}
