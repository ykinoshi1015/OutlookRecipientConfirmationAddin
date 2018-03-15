NUnit.Framework.AssertionException はユーザー コードによってハンドルされませんでした。
  HResult=-2146233088
  Message=  Expected: equivalent to < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >
  But was:  < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >
  Missing (2): < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >
  Extra (2): < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >

  Source=nunit.framework
  StackTrace:
       場所 NUnit.Framework.Assert.ReportFailure(String message)
       場所 NUnit.Framework.Assert.ReportFailure(ConstraintResult result, String message, Object[] args)
       場所 NUnit.Framework.Assert.That[TActual](TActual actual, IResolveConstraint expression, String message, Object[] args)
       場所 NUnit.Framework.CollectionAssert.AreEquivalent(IEnumerable expected, IEnumerable actual)
       場所 ORCAUnitTest.UnitTest1.GetRecipientsTest1() 場所 C:\workspaces\OutlookRecipientConfirmationAddin\sourcecode\ORCAUnitTest\UnitTest1.cs:行 352
  InnerException: 
