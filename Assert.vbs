Class ContainsTextEvaluator_
	Private rhs_
	
	Public Default Function Init(lhs)
		rhs_ = lhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (InStr(lhs,rhs_)>0)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" does not contain "+CStr(rhs_)
	End Function
End Class

Class MatchesTextEvaluator_
	Private rhs_
	
	Public Default Function Init(lhs)
		rhs_ = lhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Dim objRE: Set objRE = new RegExp
		objRE.Global = True
		objRE.IgnoreCase = True
		objRE.Pattern = rhs_
		
		Evaluate = objRE.Test(lhs)
		
		Set objRE = Nothing
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" does not match "+CStr(rhs_)
	End Function
End Class

Class StartsWithTextEvaluator_
	Private rhs_
	
	Public Default Function Init(lhs)
		rhs_ = lhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (InStr(lhs,rhs_)=1)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" does not start with "+CStr(rhs_)
	End Function
End Class

Class EndsWithTextEvaluator_
	Private rhs_
	
	Public Default Function Init(lhs)
		rhs_ = lhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Dim sCompare
		Dim lLen
   
		lLen = Len(rhs_)
		If lLen > Len(lhs) Then Exit Function
		
		sCompare = Right(lhs, lLen)
		Evaluate = (StrComp(sCompare, rhs_) = 0)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" does not end with "+CStr(rhs_)
	End Function
End Class

Class EqualToEvaluator_
	Private rhs_
	
	Public Default Function Init(lhs)
		rhs_ = lhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (lhs = rhs_)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" is not equal to "+CStr(rhs_)
	End Function
End Class

Class NaNEvaluator_
	Public Function Evaluate(lhs)
		Evaluate = (not IsNumeric(lhs))
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" is a number!"
	End Function
End Class

Class TypeOfEvaluator_
	Private rhs_
	
	Public Default Function Init(lhs)
		Set rhs_ = lhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (TypeName(lhs) = TypeName(rhs_))
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = TypeName(lhs) +" is not an instance of "+TypeName(rhs_)
	End Function
End Class

Class GreaterThanEvaluator_
	Private rhs_
	
	Public Default Function Init(rhs)
		rhs_ = rhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (lhs > rhs_)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" is not greater than "+CStr(rhs_)
	End Function
End Class

Class GreaterThanOrEqualToEvaluator_
	Private rhs_
	
	Public Default Function Init(rhs)
		rhs_ = rhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (lhs >= rhs_)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" is not greater than or equal to "+CStr(rhs_)
	End Function
End Class

Class LessThanEvaluator_
	Private rhs_
	
	Public Default Function Init(rhs)
		rhs_ = rhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (lhs < rhs_)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" is not less than "+CStr(rhs_)
	End Function
End Class

Class LessThanOrEqualToEvaluator_
	Private rhs_
	
	Public Default Function Init(rhs)
		rhs_ = rhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (lhs <= rhs_)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = CStr(lhs) +" is not less than or equal to "+CStr(rhs_)
	End Function
End Class

Class SameAsEvaluator_
	Private rhs_
	
	Public Default Function Init(rhs)
		Set rhs_ = rhs
		Set Init = Me
	End Function
	
	Public Function Evaluate(lhs)
		Evaluate = (lhs IS rhs_)
	End Function
	
	Public Function ErrorString(lhs)
		ErrorString = "Object instances are not the same"
	End Function
End Class

Class EvaluatorFactory_
	Public Function CreateEqualToEvaluator(rhs)
		Set CreateEqualToEvaluator = (New EqualToEvaluator_)(rhs)
	End Function

	Public Function CreateSameAsEvaluator(rhs)
		Set CreateSameAsEvaluator = (New SameAsEvaluator_)(rhs)
	End Function

	Public Function CreateGreaterThanEvaluator(rhs)
		Set CreateGreaterThanEvaluator = (New GreaterThanEvaluator_)(rhs)
	End Function

	Public Function CreateGreaterThanOrEqualToEvaluator(rhs)
		Set CreateGreaterThanOrEqualToEvaluator = (New GreaterThanOrEqualToEvaluator_)(rhs)
	End Function
	
	Public Function CreateLessThanEvaluator(rhs)
		Set CreateLessThanEvaluator = (New LessThanEvaluator_)(rhs)
	End Function

	Public Function CreateLessThanOrEqualToEvaluator(rhs)
		Set CreateLessThanOrEqualToEvaluator = (New LessThanOrEqualToEvaluator_)(rhs)
	End Function
	
	Public Function CreateTypeOfEvaluator(rhs)
		Set CreateTypeOfEvaluator = (New TypeOfEvaluator_)(rhs)
	End Function
	
	Public Function CreateNaNEvaluator()
		Set CreateNaNEvaluator = (New NaNEvaluator_)
	End Function
End Class

Class Be_
	Private EvaluatorFactory

	Private Sub Class_Initialize()
		Set EvaluatorFactory = new EvaluatorFactory_
	End Sub

	Private Sub Class_Terminate()
		Set EvaluatorFactory = Nothing
	End Sub

	Public Function True_()
		Set True_ = EvaluatorFactory.CreateEqualToEvaluator(True)
	End Function
	
	Public Function False_()
		False_ = False
		Set False_ = EvaluatorFactory.CreateEqualToEvaluator(False)		
	End Function
	
	Public Function Null_()
		Null_ = False
		Set Null_ = EvaluatorFactory.CreateSameAsEvaluator(Nothing)		
	End Function
	
	Public Function EqualTo(value)
		Set EqualTo = EvaluatorFactory.CreateEqualToEvaluator(value)
	End Function

	Public Function SameAs(value)
		Set SameAs = EvaluatorFactory.CreateSameAsEvaluator(value)
	End Function
	
	Public Function GreaterThan(value)
		Set GreaterThan = EvaluatorFactory.CreateGreaterThanEvaluator(value)
	End Function

	Public Function GreaterThanOrEqualTo(value)
		Set GreaterThanOrEqualTo = EvaluatorFactory.CreateGreaterThanOrEqualToEvaluator(value)
	End Function

	Public Function AtLeast(value)
		Set AtLeast = EvaluatorFactory.CreateGreaterThanOrEqualToEvaluator(value)
	End Function

	Public Function LessThan(value)
		Set LessThan = EvaluatorFactory.CreateLessThanEvaluator(value)
	End Function

	Public Function LessThanOrEqualTo(value)
		Set LessThanOrEqualTo = EvaluatorFactory.CreateLessThanOrEqualToEvaluator(value)
	End Function
	
	Public Function AtMost(value)
		Set AtMost = EvaluatorFactory.CreateLessThanOrEqualToEvaluator(value)
	End Function
	
	Public Function TypeOf_(value)
		Set TypeOf_ = EvaluatorFactory.CreateTypeOfEvaluator(value)
	End Function

	Public Function NaN()
		Set NaN = EvaluatorFactory.CreateNaNEvaluator()
	End Function
End Class

Class TextEvaluatorFactory_
	Public Function CreateStartsWithEvaluator(rhs)
		Set CreateStartsWithEvaluator = (New StartsWithTextEvaluator_)(rhs)
	End Function

	Public Function CreateEndsWithEvaluator(rhs)
		Set CreateEndsWithEvaluator = (New EndsWithTextEvaluator_)(rhs)
	End Function

	Public Function CreateMatchesEvaluator(rhs)
		Set CreateMatchesEvaluator = (New MatchesTextEvaluator_)(rhs)
	End Function
	
	Public Function CreateContainsEvaluator(rhs)
		Set CreateContainsEvaluator = (New ContainsTextEvaluator_)(rhs)
	End Function
End Class

Class Text_
	Private TextEvaluatorFactory

	Private Sub Class_Initialize()
		Set TextEvaluatorFactory = new TextEvaluatorFactory_
	End Sub

	Private Sub Class_Terminate()
		Set TextEvaluatorFactory = Nothing
	End Sub

	Public Function StartsWith(value)
		Set StartsWith = TextEvaluatorFactory.CreateStartsWithEvaluator(value)
	End Function

	Public Function EndsWith(value)
		Set EndsWith = TextEvaluatorFactory.CreateEndsWithEvaluator(value)
	End Function
	
	Public Function Matches(value)
		Set Matches = TextEvaluatorFactory.CreateMatchesEvaluator(value)
	End Function

	Public Function Contains(value)
		Set Contains = TextEvaluatorFactory.CreateContainsEvaluator(value)
	End Function
End Class

Class Assert_
	Public Sub That(expression, evaluator)
		if not evaluator.Evaluate(expression) then
			Call Err.Raise(51, "Assert Failed!", evaluator.ErrorString(expression))
		end if
		
		Set evaluator = Nothing
	End Sub
End Class
