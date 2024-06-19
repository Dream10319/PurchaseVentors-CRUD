<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="PurchaseVentor.aspx.vb" Inherits="PurchaseVentors.PurchaseVentor" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style>
        body {
            /*background-image: url("https://urldefense.com/v3/__https://img.freepik.com/free-vector/map-point-abstract-3d-polygonal-wireframe-airplane-blue-night-sky-with-dots-stars-illustration-background_587448-568.jpg*22);*/__;JSo!!Dhw9WWooB8bE!ocejLj-0m3YiODlOGt3GFrj7e-lmyWEsC-S9hBlQ4_y3hTBV2L_epirUVFtudVxzWubqwh4aSz1sBQV97EYXs9I$ 
            background-repeat: no-repeat;
            background-size: cover;
            background-position: center;
            max-width: max-content;
            margin: auto;
            margin-top: 20px;
            text-align: center;
        }

        .titles {
            font: Calibri;
            font-size: xx-large;
           /* color: #91c2c5;*/
           color:black;
           font-weight:bold;
        }

        #form1 {
            background: #FDC011;
            box-sizing: border-box;
            box-shadow: 0 15px 25px rgba(0, 0, 0, .7);
            border-radius: 10px;
        }

        .forms > td {
            border:none
        }

        .borderspace {
            border-spacing: 2em;
        }

        .lbl1 {
            background-color: beige;
        }

        .lbl2 {
            color: white;
        }

        .lbl3 {
            color: blue;

        }
        .lbl4{

         color:black;
        }
        .lbl5{
            color:black;
            border:3px solid black;
            padding:5px;
            margin:5px;


        }

        .pt {
            cursor: pointer;
            color: Blue;
        }

        .lblS {
            font: 400 130px/0.8 'Great Vibes', Helvetica, sans-serif;
            font-size: x-large;
            font-weight: 600;
            color: black;
        }

        .auto-style1 {
            margin-bottom: 0px;
            border: none;
        }

        :root {
            --arrow-bg: rgba(255, 255, 255, 0.3);
            --arrow-icon: url(https://urldefense.com/v3/__https://upload.wikimedia.org/wikipedia/commons/9/9d/Caret_down_font_awesome_whitevariation.svg__;!!Dhw9WWooB8bE!ocejLj-0m3YiODlOGt3GFrj7e-lmyWEsC-S9hBlQ4_y3hTBV2L_epirUVFtudVxzWubqwh4aSz1sBQV91CruelM$ );
            --option-bg: white;
            --select-bg: rgba(255, 255, 255, 0.2);
        }
        select {
            /* Reset */
            text-align: center;
            -webkit-appearance: none;
                -moz-appearance: none;
                  /*  appearance: none;*/
            border: 0;
            outline: 0;
            font: inherit;
            /* Personalize */
            width: 100%;
            height: 2.5rem;
            padding-right: 2.7rem;
            background: var(--arrow-icon) no-repeat right 0.8em center/1.4em, linear-gradient(to left, var(--arrow-bg) 3em, var(--select-bg) 3em);
            color: black;
            border-radius: 0.25em;
            box-shadow: 0 0 1em 0 rgba(0, 0, 0, 0.2);
            cursor: pointer;
            /* Remove IE arrow */
            /* Remove focus outline */
            /* <option> colors */
        }
        select::-ms-expand {
            display: none;
        }
        select:focus {
            outline: none;
        }
        select option {
            color: inherit;
            background-color: var(--option-bg);
        }
         .user-box1 {
            position: relative;
            font-display:inherit;
      

        }
    
        .user-box {
            position: relative;

        }
        .user-box input {
            width: 100%;
            padding: 10px 0;
            font-size: 20px;
            color: #000;
            border: none;
            border-bottom: 1px solid #000;
            outline: none;
            background: transparent;
            margin-left: 20px;
        }
        .user-box label {
            position: absolute;
            top:0;
            left: 0;
            padding: 10px 0;
            font-size: 20px;
            color: #000;
            pointer-events: none;
            transition: .5s;
            margin-left: 20px;
        }
        .user-box input:first-child ~ label,
        .user-box input:valid ~ label {
            top: -20px;
            left: 0;
            color: rgb(0,2,103);
            font-size: 20px;
        }
        .user-box input:focus ~ label {
            top: -20px;
            left: 0;
            color: rgb(0,2,103);
            font-size: 20px;
        }

        .shadowbutton {
            background-color: #003366;
            border: none;
            color: #d7d6d6;
            padding: 10px 15px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            margin: 4px 2px;
            cursor: pointer;
            -webkit-transition-duration: 0.4s; /* Safari */
            transition-duration: 0.4s;
            border-radius: 10px;
            box-shadow: 10px 5px 5px grey;

        }

        .shadowbutton:hover {
            box-shadow: 12px 7px 8px grey;
        }
    </style>
</head>
<body>
    <label class="titles" runat="server"><strong> PurchaseVendors </strong></label>
    <form id="form1" runat="server" style="padding:10px">
        <div style="text-align:left;">
            <asp:CheckBox ID="chk1" runat="server" Text ="Inactive" TextAlign="Left" AutoPostBack="true" />
        </div>
        <div>
           <asp:Table runat="server" Width="848px" Height="300px" CssClass="auto-style1">

                <asp:TableRow CssClass="forms">
                    <asp:TableCell ColumnSpan="4" Height="3em">
                        <asp:Label ID="lalDescription" runat="server" Text="Add Title" Font-Size="20"
                            Font-Underline="true" CssClass="lbl4">
                        </asp:Label>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow CssClass="forms">
                    <asp:TableCell ColumnSpan="4" Height="3em">
                        &nbsp;&nbsp;&nbsp;
                        <asp:Table runat="server">
                            <asp:TableRow>          
                                <asp:TableCell CssClass="user-box" Width="800px" >
                                   
                                    <asp:TextBox ID="txtProgramTitle" runat="server"></asp:TextBox>
                                    <label>Program Title</label>
                                </asp:TableCell>
                                <asp:TableCell Width="190px">
                                    <asp:Button ID="addTitle" runat="server"   Text="Add Title" Font-Size="14"  Font-Bold="true" CssClass="shadowbutton" OnClick="addTitle_Click1" />
                                </asp:TableCell>
                            </asp:TableRow>

                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow CssClass="forms">
                    <asp:TableCell ColumnSpan="4" Height="3em">
                        <asp:Label ID="lblReceipt" runat="server" Text="Change Title"
                            Font-Size="20" Font-Underline="true" CssClass="lbl4"></asp:Label>
                    </asp:TableCell>

                </asp:TableRow>

            <%-- I acknowledge receipt --%>

                <asp:TableRow CssClass="forms">
                    <asp:TableCell ColumnSpan="4" Height="3em">
                        <asp:Table runat="server">
                            <asp:TableRow>              
                                <asp:TableCell Width="300px">
                                    <asp:DropDownList ID="ddVendorNameUpdate" Width="300px" runat="server" Font-Size="12"
                                        CssClass="" Style ="margin-left:20px">
<%--                                        <asp:ListItem Value=""></asp:ListItem>
                                        <asp:ListItem Value="China da zhong"></asp:ListItem>
                                        <asp:ListItem Value="Toyato"></asp:ListItem>
                                        <asp:ListItem Value="USA"></asp:ListItem>
                                        <asp:ListItem Value="Germen"></asp:ListItem>--%>
                                    </asp:DropDownList>
                                </asp:TableCell>
                                <asp:TableCell CssClass="user-box" Width="300px" VerticalAlign="Bottom" >
                                    <asp:TextBox ID="txtVendorNameChg" runat="server" name=""></asp:TextBox>
                                    <label>Program Title Change</label>
                                </asp:TableCell>
                                <asp:TableCell Width="190px">
                                    <asp:Button ID="changeTitle" runat="server"
                                    Text="Change Title" Font-Size="14"
                                    Font-Bold="true" CssClass="shadowbutton" OnClick="changeTitle_Click"/>
                                </asp:TableCell>
                            </asp:TableRow>

                        </asp:Table>
                    </asp:TableCell>         
                </asp:TableRow>
            </asp:Table>  
        </div>
    </form>
</body>
</html>
