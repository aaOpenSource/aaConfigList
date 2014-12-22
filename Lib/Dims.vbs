	Dim strConnectionString, objADOConn, objADOrs0, objADOrs1, objADOrs2, objADOrs3, objADOrs4, intPosition, intSize, intLocations(), intLocation, strServer, strDB, strBaseArea, strBaseDir, strUser, blnFlag, intBound, intPack_ID, strHier_Name, strFile_Name, strFolder, strPath
    Dim strPwd, objshell, strCode, strSQL,  strValue, objArgs, objFile, objDict1, objDict2, strText, strRawData, x, y, z, intMTPosition, strEmpty, strTest, intGobject_Id, intDFGobject_Id, intTmpltDef_Id, intCtgry_Id
	Dim strFile, objHex2Str, intLevel, intPointer(), intGobject_id_0(), intGobject_id_1(), intGobject_id_2(), intGobject_id_3(), intGobject_id_4(), intGobject_id_5(), intGobject_id_6()
	Dim intGobject_id_7(), intGobject_id_8(), intGobject_id_9(), strTag_Name_0(), strTag_Name_1(), strTag_Name_2(), strTag_Name_3(), strTag_Name_4()
	Dim strTag_Name_5(), strTag_Name_6(), strTag_Name_7(), strTag_Name_8(), strTag_Name_9(), intChkd_Pack_Id, intChkd_Pack_id_0(), intChkd_Pack_id_1(), intChkd_Pack_id_2()
	Dim intChkd_Pack_id_3(), intChkd_Pack_id_4(), intChkd_Pack_id_5(), intChkd_Pack_id_6(), intChkd_Pack_id_7(), intChkd_Pack_id_8(), intChkd_Pack_id_9()
	Dim intDeploy_Ver, intDeploy_Ver_0(), intDeploy_Ver_1(), intDeploy_Ver_2(), intDeploy_Ver_3(), intDeploy_Ver_4(), intDeploy_Ver_5(), sngSecs, sngFirstSecs
	Dim intDeploy_Ver_6(), intDeploy_Ver_7(), intDeploy_Ver_8(), intDeploy_Ver_9(), intDep_Pack_Id, intDep_Pack_id_0(), intDep_Pack_id_1()
	Dim intConfig_Ver, intConfig_Ver_0(), intConfig_Ver_1(), intConfig_Ver_2(), intConfig_Ver_3(), intConfig_Ver_4(), intConfig_Ver_5(), blnNightCheck
	Dim intConfig_Ver_6(), intConfig_Ver_7(), intConfig_Ver_8(), intConfig_Ver_9(), intDep_Pack_id_2(), intDep_Pack_id_3(), intDep_Pack_id_4()
	Dim intDep_Pack_id_5(), intDep_Pack_id_6(), intDep_Pack_id_7(), intDep_Pack_id_8(), intDep_Pack_id_9(), strArea, strYear, strMonth, strDay
	Dim strScratch, intScratch, intValue, dblValue, lngObjCnt, lngAttrCnt, datStart, datEnd, lngDateBias, lngDateBiasKey, lngActiveTimeBias, lngActiveTimeBiasKey
    Dim strRun, strMenu, strStartDir, strWorkDir, strUDA, strOptArgs, strHour, strMinute, blnAttrib, intDFGobject_id_0(), intDFGobject_id_1(), intDFGobject_id_2()
	Dim intDFGobject_id_3(), intDFGobject_id_4(), intDFGobject_id_5(), intDFGobject_id_6(), intDFGobject_id_7(), intDFGobject_id_8(), intDFGobject_id_9()
	Dim intTmpltDef_id_0(), intTmpltDef_id_1(), intTmpltDef_id_2(), intTmpltDef_id_3(), intTmpltDef_id_4(), intTmpltDef_id_5(), intTmpltDef_id_6()
	Dim intTmpltDef_id_7(), intTmpltDef_id_8(), intTmpltDef_id_9(), intCtgry_id_0(), intCtgry_id_1(), intCtgry_id_2(), intCtgry_id_3(), intCtgry_id_4()
	Dim intCtgry_id_5(), intCtgry_id_6(), intCtgry_id_7(), intCtgry_id_8(), intCtgry_id_9(), strItem, strAlias, strScanGroup, objADOStream, bytStream