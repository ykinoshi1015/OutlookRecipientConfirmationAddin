NUnit.Framework.AssertionException �̓��[�U�[ �R�[�h�ɂ���ăn���h������܂���ł����B
  HResult=-2146233088
  Message=  Expected: equivalent to < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >
  But was:  < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >
  Missing (2): < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >
  Extra (2): < <Castle.Proxies.RecipientProxy>, <Castle.Proxies.RecipientProxy> >

  Source=nunit.framework
  StackTrace:
       �ꏊ NUnit.Framework.Assert.ReportFailure(String message)
       �ꏊ NUnit.Framework.Assert.ReportFailure(ConstraintResult result, String message, Object[] args)
       �ꏊ NUnit.Framework.Assert.That[TActual](TActual actual, IResolveConstraint expression, String message, Object[] args)
       �ꏊ NUnit.Framework.CollectionAssert.AreEquivalent(IEnumerable expected, IEnumerable actual)
       �ꏊ ORCAUnitTest.UnitTest1.GetRecipientsTest1() �ꏊ C:\workspaces\OutlookRecipientConfirmationAddin\sourcecode\ORCAUnitTest\UnitTest1.cs:�s 352
  InnerException: 
