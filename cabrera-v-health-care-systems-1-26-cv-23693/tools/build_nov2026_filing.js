const fs=require('fs');
const d=require('/opt/npm-tools/node_modules/docx');
const {Document,Packer,Paragraph,TextRun,AlignmentType,ImageRun,Table,TableRow,TableCell,WidthType,BorderStyle,Footer,PageNumber}=d;
const sig=fs.readFileSync('/home/claude/sig_cabrera_clean.png');
const F="Times New Roman", DATE="November 6, 2026";
const t=(s,o={})=>new TextRun({text:s,font:F,size:24,...o});
const body=(runs,o={})=>new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{line:480,after:0},indent:{firstLine:720},children:Array.isArray(runs)?runs:[t(runs)],...o});
const head=(s)=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{line:480,after:0},keepNext:true,children:[t(s,{bold:true})]});
const sub=(s)=>new Paragraph({alignment:AlignmentType.LEFT,spacing:{line:240,before:120,after:240},keepNext:true,indent:{left:720,hanging:720},children:[t(s,{bold:true})]});
const single=(s,o={})=>new Paragraph({spacing:{line:240,after:0},alignment:o.al||AlignmentType.LEFT,children:[t(s,o.r||{})],...(o.p||{})});
const blank=()=>single("");
const nb={style:BorderStyle.NONE,size:0,color:"FFFFFF"};
const line={style:BorderStyle.SINGLE,size:6,color:"000000"};
function caption(title){
 const L=[single("RIGOBERTO CABRERA,"),blank(),single("          Plaintiff,"),blank(),single("v."),blank(),
  single("HEALTHCARE SYSTEMS OF AMERICA-FLORIDA LLC, d/b/a PALMETTO GENERAL HOSPITAL; HSA PGH, LLC, d/b/a PALMETTO GENERAL HOSPITAL; and JOHN DOE, an unidentified MRI Technologist,"),blank(),single("          Defendants."),blank()];
 const R=[single("Case No.: 1:26-cv-23693-MFE")];
 const tbl=new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[5200,4160],rows:[new TableRow({children:[
  new TableCell({width:{size:5200,type:WidthType.DXA},borders:{top:nb,left:nb,bottom:line,right:line},margins:{right:120},children:L}),
  new TableCell({width:{size:4160,type:WidthType.DXA},borders:{top:nb,left:nb,bottom:nb,right:nb},margins:{left:300,top:600},children:R})]})]});
 return [single("UNITED STATES DISTRICT COURT",{al:AlignmentType.CENTER,r:{bold:true}}),single("SOUTHERN DISTRICT OF FLORIDA",{al:AlignmentType.CENTER,r:{bold:true}}),blank(),tbl,blank(),
  ...title.map(x=>single(x,{al:AlignmentType.CENTER,r:{bold:true}})),blank()];
}
function sigblock(label){
 return [blank(),
  new Paragraph({spacing:{line:240,before:120},keepNext:true,children:[t(`Respectfully submitted this 6th day of November, 2026.`)]}),
  new Paragraph({spacing:{line:240,before:480},keepNext:true,indent:{left:4320},border:{bottom:{style:BorderStyle.SINGLE,size:6,color:"000000",space:1}},children:[new ImageRun({type:"png",data:sig,transformation:{width:190,height:34},altText:{title:"sig",description:"sig",name:"sig"}})]}),
  ...["RIGOBERTO CABRERA","Plaintiff, Pro Se","Reg. No. 33513-018","Federal Correctional Institution Miami","P.O. Box 779800","Miami, Florida 33177"].map((s,i)=>new Paragraph({spacing:{line:240,before:i==0?80:0},indent:{left:4320},keepNext:i<5,children:[t(s,{bold:i==0})]}))];
}
const footer=new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({children:[PageNumber.CURRENT],font:F,size:24})]})]});
const page={size:{width:12240,height:15840},margin:{top:1440,bottom:1440,left:1440,right:1440}};
const mk=(children,ft=true)=>new Document({styles:{default:{document:{run:{font:F,size:24}}}},sections:[{properties:{page},footers:ft?{default:footer}:undefined,children}]});

// ---------- MOTION ----------
const M=[...caption(["PLAINTIFF’S MOTION FOR LEAVE TO ACCEPT LATE-FILED RESPONSE","TO ORDER TO SHOW CAUSE, AND REQUEST FOR A RULING ON HIS","MOTION TO PROCEED IN FORMA PAUPERIS AND FOR SERVICE","OF PROCESS BY THE UNITED STATES MARSHAL"]),
body("Plaintiff Rigoberto Cabrera, appearing pro se, respectfully asks the Court to do three things. First, accept his Response to the Court’s Order to Show Cause, which the Clerk docketed on July 1, 2026, nine days after the June 22, 2026 deadline. Second, rule on his Motion for Leave to Proceed in Forma Pauperis (DE 3), which has been pending since he filed this case on May 22, 2026. Third, direct the United States Marshal to serve process on the Defendants. Plaintiff is incarcerated and cannot move this case forward without the Court’s help."),
head("BACKGROUND"),
body("Plaintiff is a federal prisoner at the Federal Correctional Institution in Miami. On May 22, 2026, he filed this action for battery under Florida law, invoking diversity jurisdiction under 28 U.S.C. § 1332(a)(2) because he is a citizen of a foreign state. He filed a motion for leave to proceed in forma pauperis (DE 3) with his complaint and later filed a supplement to it (DE 6)."),
body("On June 1, 2026, the Court issued an Order to Show Cause (DE 4) directing Plaintiff to demonstrate that the Court has subject-matter jurisdiction under 28 U.S.C. § 1332, with a response due June 22, 2026. Plaintiff prepared a Response, a First Amended Complaint naming the Defendants’ corporate structure, a Declaration, and supporting public records. Because he is incarcerated, he has no access to the Court’s electronic filing system and must send and receive his legal mail by regular mail through a person outside the facility. Part of the supporting material was a defendants’ answer that Healthcare Systems of America-Florida LLC and related entities filed in another case in this District on June 11, 2026, eleven days before the deadline. The package was sent out for mailing on June 25, 2026, and the Clerk docketed it on July 1, 2026."),
body("No Defendant has been served or has appeared. The Court has not ruled on the Response, on the motion to proceed in forma pauperis, or on service. The docket has been quiet since July 1, 2026."),
head("LEGAL STANDARD"),
body([t("When a deadline has passed, the Court “may, for good cause, extend the time . . . on motion made after the time has expired if the party failed to act because of excusable neglect.” Fed. R. Civ. P. 6(b)(1)(B). The Supreme Court has explained that “the determination is at bottom an equitable one, taking account of all relevant circumstances surrounding the party’s omission.” "),t("Pioneer Inv. Servs. Co. v. Brunswick Assocs. Ltd. P’ship",{italics:true}),t(", 507 U.S. 380, 395 (1993). The relevant circumstances include “the danger of prejudice to the [opposing party], the length of the delay and its potential impact on judicial proceedings, the reason for the delay, including whether it was within the reasonable control of the movant, and whether the movant acted in good faith.” "),t("Id.",{italics:true}),t(" Each factor favors Plaintiff.")]),
head("ARGUMENT"),
sub("I.\tThe Court Should Accept Plaintiff’s Response Because the Short Delay Was Excusable Neglect."),
body([t("No one is harmed by accepting the Response. ",{bold:true}),t("No Defendant has been served or has appeared, so no one has relied on the missed date or lost anything by it. The Court has set no hearing, trial date, or other deadline that the late Response disturbed. Accepting it costs the Court nothing and lets the Court decide jurisdiction on a full record.")]),
body([t("The delay was brief. ",{bold:true}),t("The Response was docketed nine days after the deadline. In the months since, the case has not moved, so the delay did not slow anything the Court was doing.")]),
body([t("The reason for the delay is understandable. ",{bold:true}),t("Plaintiff is a prisoner with no electronic filing access. To answer the Order, he had to assemble a response and public records from other courts, including a filing made only on June 11, 2026, and then get the package into the mail through someone outside the prison. Those steps ran past the deadline. Plaintiff did not ignore the Court’s Order. He answered it as fully as he could, and he is not asking the Court to excuse a refusal to comply. He acknowledges that he should have asked for more time before June 22, and he asks the Court to treat his omission as the kind of neglect "),t("Pioneer",{italics:true}),t(" allows the Court to excuse.")]),
body([t("Plaintiff acted in good faith. ",{bold:true}),t("The Response and First Amended Complaint were prepared to answer the Court’s concern directly, with documents the Court can verify. Plaintiff now comes forward on his own to explain the delay and ask for leave.")]),
sub("II.\tThe Court Should Rule on Plaintiff’s Pending Motion to Proceed in Forma Pauperis (DE 3)."),
body("Plaintiff has no money to pay the filing fee. His motion and supplement tell the Court so, and his inmate trust account is nearly empty. Because he cannot pay to serve the Defendants himself, the motion is the key that unlocks the rest of the case. It has now been pending for more than five months. Plaintiff respectfully asks the Court to decide it."),
sub("III.\tThe Court Should Direct the United States Marshal to Serve Process."),
body([t("When a plaintiff is allowed to proceed in forma pauperis, Congress directs that the officers of the court issue and serve all process. 28 U.S.C. § 1915(d). The Federal Rules say the same thing: the Court “may order that service be made by a United States marshal,” and it “must so order if the plaintiff is authorized to proceed in forma pauperis under 28 U.S.C. § 1915.” Fed. R. Civ. P. 4(c)(3). Plaintiff is in custody, has no means to pay a process server, and cannot serve anyone himself. If the Court grants the motion to proceed in forma pauperis, Plaintiff asks the Court to order the Marshal to serve the summons and First Amended Complaint on each Defendant, to tell Plaintiff what, if anything, he must provide for that purpose, and to confirm that the time to complete service runs from the date of the Court’s order.")]),
sub("IV.\tPlaintiff Is Filing a Separate Request for Judicial Notice."),
body("Plaintiff prepared a Request for Judicial Notice of public court records when he filed his Response, but it does not appear on the docket. He is filing it again with this Motion so the Court has it."),
head("CERTIFICATE OF CONFERRAL"),
body("No Defendant has been served or has appeared, so there is no opposing party with whom Plaintiff could confer before filing this Motion."),
head("CONCLUSION"),
body("For these reasons, Plaintiff respectfully asks the Court to: (a) grant leave under Federal Rule of Civil Procedure 6(b)(1)(B) and accept his Response to the Order to Show Cause as timely filed; (b) rule on his Motion for Leave to Proceed in Forma Pauperis (DE 3); (c) if that motion is granted, order the United States Marshal to serve the summons and First Amended Complaint on each Defendant under 28 U.S.C. § 1915(d) and Federal Rule of Civil Procedure 4(c)(3), and provide that the time for service runs from that order; and (d) grant any further relief the Court finds just."),
...sigblock()];

// ---------- RJN ----------
const it=(x)=>t(x,{italics:true});
const R=[...caption(["PLAINTIFF’S REQUEST FOR JUDICIAL NOTICE","IN SUPPORT OF PLAINTIFF’S RESPONSE TO ORDER TO SHOW CAUSE"]),
body("Under Federal Rule of Evidence 201(b) and 201(c)(2), Plaintiff Rigoberto Cabrera, appearing pro se, asks the Court to take judicial notice of the five public records identified below. Two of them, the complaint and the answer in the Maxfield case, are attached to this Request as Exhibits A and B. The other three are government and court records whose contents are identified specifically below. Plaintiff first prepared this Request with his Response to the Order to Show Cause, but it does not appear on the docket, so he is filing it again."),
head("LEGAL STANDARD"),
body("A court “must take judicial notice if a party requests it and the court is supplied with the necessary information.” Fed. R. Evid. 201(c)(2). The Court may notice a fact that is not subject to reasonable dispute because it “can be accurately and readily determined from sources whose accuracy cannot reasonably be questioned.” Fed. R. Evid. 201(b)(2). It may do so at any stage of the proceeding. Fed. R. Evid. 201(d). Plaintiff asks the Court to notice that each record was filed and what each says. He does not ask the Court to accept the truth of any disputed statement in them."),
head("RECORDS"),
sub("I.\tFlorida Department of State — Entity Record for Healthcare Systems of America–Florida LLC"),
body("Fla. Dep’t of State, Div. of Corps., Entity Records for Healthcare Systems of America–Florida LLC, Doc. No. M24000012582 (filed Oct. 2, 2024), https://search.sunbiz.org (last visited June 22, 2026). Plaintiff asks the Court to notice that this record shows Healthcare Systems of America–Florida LLC registered with the Florida Division of Corporations as a Nevada limited liability company (state of organization: NV), filed October 2, 2024, with a principal address of 505 North Brand Boulevard, Suite 1200, Glendale, California 91203, and a registered agent address in Miami, Florida. It is a record of a state agency kept in an official searchable database."),
sub("II.\tComplaint, Sarian v. Paronyan, No. 26STCV05675 (Cal. Super. Ct., L.A. Cnty.)"),
body("Complaint, Sarian v. Paronyan, No. 26STCV05675 (Cal. Super. Ct., L.A. Cnty. Feb. 24, 2026) (Quinn Emanuel Urquhart & Sullivan, LLP for Plaintiffs). Plaintiff asks the Court to notice that (1) this complaint was filed in the Superior Court of California, County of Los Angeles, on February 24, 2026; (2) Paragraph 10 states that “Plaintiff Mike Sarian is an individual residing in Los Angeles County, California”; (3) Paragraph 13 states that the Sarian Family Trust is a “California trust with its principal place of administration in Los Angeles County” and holds “1,000 shares (100%) of Plaintiff Healthcare Systems of America Corp.”; and (4) Paragraph 15 states that Healthcare Systems of America Corp. “is a Nevada corporation headquartered in Glendale, California,” “owns and operates eight hospitals outside California,” and “is wholly owned by the Sarian Family Trust.” These are statements in a public court filing."),
sub("III.\tDeclaration of Mike Sarian, Sarian v. Paronyan, No. 26STCV05675 (Cal. Super. Ct., L.A. Cnty.)"),
body("Declaration of Mike Sarian in Support of Plaintiffs’ Ex Parte Application for a Temporary Restraining Order and Order to Show Cause Re: Preliminary Injunction ¶¶ 4, 6, 8, 11, 20, Sarian v. Paronyan, No. 26STCV05675 (Cal. Super. Ct., L.A. Cnty. Feb. 24, 2026). Plaintiff asks the Court to notice that (1) this declaration was filed under penalty of perjury in Los Angeles County Superior Court on February 24, 2026; (2) Paragraph 4 states, “I formed Healthcare Systems of America Corp. (‘HSA’) in September 2024 to pursue additional hospital acquisitions”; (3) Paragraph 6 states, “I am the CEO of all three companies and hold majority or sole ownership of each: 51% of AHS, 100% of HSA, and 60% of NOR”; (4) Paragraph 8 states, “Aramais Paronyan maintains minority ownership stakes in AHS (49%) and NOR (40%), and holds no ownership interest in HSA”; (5) Paragraph 11 states, “He [Gill] does not own shares in any of the companies”; and (6) Paragraph 20 states, “Mr. Paronyan has never had any ownership interest or involvement in HSA.” These are statements in a public court filing."),
sub("IV.\tComplaint, Maxfield v. Healthcare Systems of America–Florida LLC, No. 0:26-cv-61343-MD (Exhibit A)"),
body("Complaint, Maxfield v. Healthcare Sys. of Am.–Fla. LLC, No. 0:26-cv-61343-MD (S.D. Fla. May 5, 2026) (Doc. 1) (Perera Law Group, P.A. for Plaintiff). Plaintiff asks the Court to notice that Paragraph 7 states, “HSAF is a foreign limited liability company registered to conduct business throughout the State of Florida,” and that Paragraphs 5 and 6 describe co-defendants HSA FMC LLC and HSA NSMC LLC as Florida limited liability companies organized under Florida law. The difference between Paragraphs 5–6 and Paragraph 7 in the same pleading can be determined from this Court’s own records."),
sub("V.\tAnswer and Affirmative Defenses, Maxfield v. Healthcare Systems of America–Florida LLC, No. 0:26-cv-61343-MD (Exhibit B)"),
body("Answer ¶ 7, Maxfield v. Healthcare Sys. of Am.–Fla. LLC, No. 0:26-cv-61343-MD (S.D. Fla. June 11, 2026) (Doc. 14). Plaintiff asks the Court to notice that, in this answer filed by the defendants’ own counsel, Paragraph 7 states, “Defendants admit that HSAF is a limited liability company registered to do business in Florida.” This is a statement in a public court filing in a live proceeding in this District."),
head("CONCLUSION"),
body("Plaintiff respectfully asks the Court to take judicial notice of the five records above for the limited purposes stated, in connection with Plaintiff’s Response to the Order to Show Cause and the First Amended Complaint."),
...sigblock()];

// ---------- COVER LETTER ----------
const C=[
 single("Rigoberto Cabrera, Reg. No. 33513-018"),single("Federal Correctional Institution Miami"),single("P.O. Box 779800"),single("Miami, Florida 33177"),blank(),
 single(DATE),blank(),
 single("Clerk of Court"),single("United States District Court, Southern District of Florida"),single("Wilkie D. Ferguson, Jr. United States Courthouse"),single("400 North Miami Avenue"),single("Miami, Florida 33128"),blank(),
 new Paragraph({spacing:{line:240,after:0},children:[t("Re:\t",{bold:true}),t("Cabrera v. Healthcare Systems of America–Florida LLC, et al.",{bold:true,italics:true}),t(", Case No. 1:26-cv-23693-MFE (Magistrate Judge Elfenbein)",{bold:true})]}),blank(),
 single("Dear Clerk:"),blank(),
 new Paragraph({alignment:AlignmentType.LEFT,spacing:{line:276,after:200},children:[t("I am the Plaintiff in this case. Please file the two enclosed documents on the docket and bring them to the attention of the Court:")]}),
 new Paragraph({spacing:{line:276,after:80},indent:{left:720,hanging:360},children:[t("1.\tPlaintiff’s Motion for Leave to Accept Late-Filed Response to Order to Show Cause, and Request for a Ruling on His Motion to Proceed in Forma Pauperis and for Service of Process by the United States Marshal; and")]}),
 new Paragraph({spacing:{line:276,after:200},indent:{left:720,hanging:360},children:[t("2.\tPlaintiff’s Request for Judicial Notice, with Exhibits A and B.")]}),
 new Paragraph({spacing:{line:276,after:200},children:[t("I am incarcerated and cannot file electronically, so I am sending these by regular mail. Thank you for your help.")]}),
 single("Respectfully,"),
 new Paragraph({spacing:{line:240,before:360},indent:{right:5400},border:{bottom:{style:BorderStyle.SINGLE,size:6,color:"000000",space:1}},children:[new ImageRun({type:"png",data:sig,transformation:{width:190,height:34},altText:{title:"sig",description:"sig",name:"sig"}})]}),
 new Paragraph({spacing:{line:240,before:80},children:[t("Rigoberto Cabrera, Plaintiff, Pro Se")]}),blank(),single("Enclosures (2)")];

(async()=>{
 fs.writeFileSync('Cabrera_Motion_Late_Response_IFP_Marshal.docx',await Packer.toBuffer(mk(M)));
 fs.writeFileSync('Cabrera_Request_for_Judicial_Notice.docx',await Packer.toBuffer(mk(R)));
 fs.writeFileSync('Cabrera_Cover_Letter_Clerk.docx',await Packer.toBuffer(mk(C,false)));
})();
